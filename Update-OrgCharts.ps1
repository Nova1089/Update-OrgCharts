# functions
function Initialize-HeaderNames # remember to tweak this function if your spreadsheet uses different header names
{
    $script:upn = "Primary Email"
    $script:jobTitle = "Jobs (HR)(1)"
    $script:department = "Department"
    $script:managerUPN = "Manager Email"
}

function Show-Introduction
{
    Write-Host ("This script updates job title, department, and manager for a provided spreadsheet of users.`n" +
    "Please use a spreadsheet that has headers with these exact names:`n" +
    "$upn, $jobTitle, $department, $managerUPN`n`n" +
    "NOTE: It's okay if your spreadsheet has other headers too and they need be in no particular order.`n") -ForegroundColor DarkCyan

    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule($moduleName)
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor DarkCyan
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor DarkCyan
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "(?<!\S)y(?!\S)") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
        "1. Open Powershell as admin.`n" +
        "2. CD into script directory.`n" +
        "3. Run .\scriptname`n") -ForegroundColor Red
        exit
    }
}

function TryConnect-AzureAD
{
    $connectedToAzureAD = Test-ConnectedToAzureAD
    while (-not($connectedToAzureAD))
    {
        Write-Host "Connecting to Azure AD..." -ForegroundColor DarkCyan
        Connect-AzureAD -ErrorAction SilentlyContinue | Out-Null
        $connectedToAzureAD = Test-ConnectedToAzureAD

        if (-not($connectedToAzureAD))
        {
            Read-Host -Prompt "Failed to connect to Azure AD. Press Enter to try again"
        }
    }
}

function Test-ConnectedToAzureAD
{
    try
    {
        Get-AzureADTenantDetail
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException]
    {
        return $false
    }
    return $true
}

function TryGet-UserCSV
{
    do
    {
        $csvPath = Prompt-UserCSVPath
        $userImport = Import-CSV -Path $csvPath
    }
    while (-not(Test-CSVHasCorrectHeaders $userImport))

    Write-Host "CSV was found and validated.`n" -ForegroundColor Green

    return $userImport
}

function Prompt-UserCSVPath
{
    do
    {
        $csvPath = Read-Host "Enter path to user CSV (i.e. C:\UserExport.csv)"
        $csvPath = $csvPath.Trim('"') # trim quotes from the path if they were entered

        if ($csvPath -notlike "*.csv")
        {
            Write-Warning "Please enter path to a CSV file."
            continue
        }

        $csvFile = Get-Item -Path $csvPath -ErrorAction SilentlyContinue
        if ($null -eq $csvFile)
        {
            Write-Warning "File not found."
        }
    }
    while ($null -eq $csvFile)    
    return $csvPath
}

function Test-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name $upn))
    {
        Write-Warning "This CSV file is missing a header called $upn to represent the user principal name."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name $jobTitle))
    {
        Write-Warning "This CSV file is missing a header called $jobTitle to represent the user's job title."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name $department))
    {
        Write-Warning "This CSV file is missing a header called $department to represent the user's department."
        $validCSV = $false
    }
    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name $managerUPN))
    {
        Write-Warning "This CSV file is missing a header called $managerUPN to represent the manager's user principal name."
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Read-Host "Please make corrections to CSV and press Enter to continue" | Out-Null
    }

    return $validCSV
}

function Update-OrgCharts($userImport)
{
    Write-Host "Updating org charts..." -ForegroundColor DarkCyan
    $timeStamp = New-TimeStamp
    $logFilePath = "$PSScriptRoot\OrgChartChanges $timeStamp.csv"
    $totalRecordsProcessed = 0
    $totalUsersChanged = 0
    foreach ($user in $userImport)
    {        
        $totalRecordsProcessed++
        Write-Progress -Activity "Updating users with new info..." -Status "$totalRecordsProcessed records processed."

        $userUPN = $user.$upn.Trim()
        $userManagerUPN = $user.$managerUPN.Trim()

        if ($userUPN -notlike "*@blueravensolar.com") { continue } # do nothing if user doesn't have a Blue Raven email

        $foundDifference = Log-Differences -importedUser $user -logFilePath $logFilePath
        if (-not($foundDifference)) { continue } # don't try to make any changes if there are no changes to make

        if ("" -ne $user.$department)
        {
            Set-AzureADUser -ObjectID $userUPN -Department $user.$department
        }

        if ("" -ne $user.$jobTitle)
        {
            Set-AzureADUser -ObjectID $userUPN -JobTitle $user.$jobTitle
        }
        
        if ("" -ne $userManagerUPN)
        {            
            if ($userManagerUPN -notlike "*@blueravensolar.com") { continue }
            
            $manager = Get-AzureADUser -ObjectID $userManagerUPN -ErrorAction SilentlyContinue
            if ($null -eq $manager)
            {
                Write-Warning ("Manager not found.`n" +
                    "User: $($userUPN)`n" +
                    "Manager: $($userManagerUPN)")
                continue
            }
            Set-AzureADUserManager -ObjectID $userUPN -RefObjectID $manager.ObjectID
        }
        $totalUsersChanged++
    }
    Write-Host ("`nFinished updating.`n" +
        "$totalUsersChanged total users changed.`n" +
        "Changes logged in $logFilePath (if there were any).`n" +
        "Please allow a few minutes for the changes to reflect in the Microsoft 365 admin portal.`n") -ForegroundColor Green
}

function New-TimeStamp
{
    return (Get-Date -Format yyyy-MM-dd-hh-mm).ToString()
}

function Log-Differences($importedUser, $logFilePath)
{
    if ($null -eq $importedUser.$upn) { return $false }
    
    $importedUserUPN = $importedUser.$upn.Trim()
    $importedUserManagerUPN = $importedUser.$managerUPN.Trim()
    $foundDifference = $false # initializing to false prevents strange errors later where this variable is null but evaluated as true

    try # use try-catch here because -ErrorAction SilentlyContinue is not working for the Get-AzureADUser cmdlet
    {
        $currentUser = Get-AzureADUser -ObjectID $importedUserUPN -ErrorAction SilentlyContinue
    }
    catch
    {
        Write-Warning "User was not found: $($importedUserUPN)"
        return $false
    }

    if ($importedUser.$department -eq "")
    {
        Write-Warning "Provided department is blank for user $($importedUserUPN). Department was not changed."
    }
    elseif ($importedUser.$department -ne $currentUser.Department)
    {
        $foundDifference = $true
        $differences = New-DifferencesRecord
        $differences."Department Before" = $currentUser.Department
        $differences."Department After" = $importedUser.$department 
    }
    
    if ($importedUser.$jobTitle -eq "")
    {
        Write-Warning "Provided title is blank for user $($importedUserUPN). Title was not changed."
    }
    elseif ($importedUser.$jobTitle -ne $currentUser.JobTitle)
    {
        $foundDifference = $true
        if ($null -eq $differences) { $differences = New-DifferencesRecord }
        $differences."Title Before" = $currentUser.JobTitle
        $differences."Title After" = $importedUser.$jobTitle
    }

    $currentManager = Get-AzureADUserManager -ObjectID $importedUserUPN -ErrorAction SilentlyContinue

    if ($importedUserManagerUPN -eq "")
    {
        Write-Warning "Provided manager is blank for user $($importedUserUPN). Manager was not changed."
    }
    elseif ($importedUserManagerUPN -notlike "*@blueravensolar.com") 
    { 
        Write-Warning ("Provided manager email is not corporate. Manager was not changed.`n" +
            "         User: $($user.$upn)`n" + 
            "         Manager: $($user.$managerUPN)")
    }
    elseif ($importedUserManagerUPN -ne $currentManager.UserPrincipalName)
    {
        $foundDifference = $true
        if ($null -eq $differences) { $differences = New-DifferencesRecord }
        $differences."Manager Before" = $currentManager.UserPrincipalName
        $differences."Manager After" = $importedUserManagerUPN
    }

    if ($foundDifference)
    {
        $differences."UPN" = $importedUserUPN
        $differences | Export-CSV -Path $logFilePath -Append -NoTypeInformation
    }
    return $foundDifference
}

function New-DifferencesRecord
{
    return [PSCustomObject]@{
        "UPN"               = ""
        "Department Before" = ""
        "Department After"  = ""
        "Title Before"      = ""
        "Title After"       = ""
        "Manager Before"    = ""
        "Manager After"     = ""
    }
}

# main
Initialize-HeaderNames
Show-Introduction
Use-Module "AzureAD"
TryConnect-AzureAD
$userImport = TryGet-UserCSV
Read-Host "Press Enter to make the changes"
Update-OrgCharts $userImport
Read-Host "Press Enter to exit"