#Set Execution Policy 
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

#Get Product Release IDs
((Get-O365Setting).ProductReleaseIds -split ',') -match 'Access' | ForEach-Object { Remove-O365ProductReleaseId -Name $_ }

#Remove from O365 Settings
Remove-O365Setting -Name "O365ProPlusRetail.ExcludedApps" -Value 'access'

#Close All O365 Apllications that are running

#Repair O365 Silent Configuration in SCCM
