# Gather all local users
$LocalUsers = Get-LocalUser

# Gather all local groups
$LocalGroups = Get-LocalGroup

# Create an empty array to store permissions data
$PermissionsTable = @()

# Loop through each user and gather permissions
foreach ($User in $LocalUsers) {
    $UserPermissions = @{
        Name = $User.Name
        ObjectType = "User"
        Groups = (Get-LocalGroupMember -Member $User.Name | ForEach-Object { $_.Name }) -join ", "
    }
    $PermissionsTable += New-Object PSObject -Property $UserPermissions
}

# Loop through each group and gather permissions
foreach ($Group in $LocalGroups) {
    $GroupPermissions = @{
        Name = $Group.Name
        ObjectType = "Group"
        Members = (Get-LocalGroupMember $Group.Name | ForEach-Object { $_.Name }) -join ", "
    }
    $PermissionsTable += New-Object PSObject -Property $GroupPermissions
}

# Display the results in a table format
#$PermissionsTable | Format-Table -AutoSize
$PermissionsTable | Export-Csv -Path "PermissionsReport.csv" -NoTypeInformation
