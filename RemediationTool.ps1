#Remove User from all APP V AD Groups
$user = $env:USERNAME
$groups = Get-ADGroup -Filter {Name -like "AppV - Microsoft"}

foreach ($group in $groups) {
    try {
        Remove-ADPrincipalGroupMembership $user.samaccountname -member $group -confirm:$false -ErrorAction Stop
    }
    catch {
        write-warning "$_ Error removing user $($user.samaccountname)"
    }
}

#Remove Packages from APP V Commander


####Set Loop that goes through Access, Project, Visio
$remApps = @('ACCESS', 'PROJECT', 'VISIO')

foreach ($app in $remApps) {
    if ($app -eq 'ACCESS') {
        #Set Execution Policy 
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Start-Sleep -s 5

        #Get Product Release IDs
        ((Get-O365Setting).ProductReleaseIds -split ',') -match 'Access' | ForEach-Object { Remove-O365ProductReleaseId -Name $_ }
        Start-Sleep -s 5

        #Remove from O365 Settings
        Remove-O365Setting -Name "O365ProPlusRetail.ExcludedApps" -Value 'access'
        Start-Sleep -s 5
    }
    elseif ($app -eq 'PROJECT') {
        #Set Execution Policy
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Start-Sleep -s 5

        #Get Product release IDs
        ((Get-O365Setting).ProductReleaseIds -split ',') -match 'Project' | ForEach-Object { Remove-O365ProductReleaseId -Name $_ }
        Start-Sleep -s 5

        #Add O365 Settings
        Add-O365ProductReleaseId -Name (ConvertTo-ClickToRunProductReleaseId -Name 'Project Standard 2016 (volume licensed)')
        Start-Sleep -s 5

    }
    else {
        #Set Execution Policy
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Start-Sleep -s 5

        #Get Product Release IDs
        ((Get-O365Setting).ProductReleaseIds -split ',') -match 'Visio' | ForEach-Object { Remove-O365ProductReleaseId -Name $_ }
        Start-Sleep -s 5

        #Add O365 Settings
        Add-O365ProductReleaseId -Name (ConvertTo-ClickToRunProductReleaseId -Name 'Visio Standard 2016 (volume licensed)')
        Start-Sleep -s 5
    }
}


#Close All O365 Apllications that are running
#Create Array of O365 Processes
$listOfApps = @('MSACCESS','EXCEL', 'ONENOTE', 'ONENOTEM', 'OUTLOOK', 'POWERPNT', 'MSPUB', 'WINWORD','Teams')

#Loop through processes and close them
foreach ($app in $listOfApps) {
    $appName = Get-Process $app -ErrorAction SilentlyContinue
    if ($appname) {

        $appName.CloseMainWindow()
        Start-Sleep 5
        
        if (!$appName.HasExited) {
            $appName | Stop-Process -Force
        }

    }
}

#Repair O365 Silent Configuration in SCCM
