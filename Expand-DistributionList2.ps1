 
Import-Module ActiveDirectory 

Param (
    [Parameter(Mandatory=$true)][string]$Group
) 

$MasterList = @(); 
$members = Get-ADGroupMember -Identity $Group -Recursive

foreach ($member in $members) { 
    if ((get-recipient $member.name).recipienttype -eq "UserMailbox") { 
        $mbox = get-mailbox $member.SamAccountName; 
        
        $row = New-Object PSObject;  
        $row | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $mbox.PrimarySMTPAddress; 
        $row | Add-Member -MemberType NoteProperty -Name "BadItemCount" -Value "0"; 
        $row | Add-Member -MemberType NoteProperty -Name "StaffNumber" -Value $mbox.SamAccountName; 
        $row | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mbox.DisplayName; 
        $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinGB" -Value "0"; 
        $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinMB" -Value "0"; 
        $MasterList += $row 
    } else {
        $mbox = get-RemoteMailbox $member.SamAccountName; 
        $row = New-Object PSObject;  
        $row | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value "MIGRATED: $($mbox.PrimarySMTPAddress)"; 
        $row | Add-Member -MemberType NoteProperty -Name "BadItemCount" -Value "0"; 
        $row | Add-Member -MemberType NoteProperty -Name "StaffNumber" -Value $mbox.SamAccountName; 
        $row | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mbox.DisplayName; 
        $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinGB" -Value "0"; 
        $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinMB" -Value "0"; 
        $MasterList += $row 
    }
} 

Write-Host "Total recursive uniquie members: $($MasterList.length)" 
$MasterList | Export-csv -path ".\$($Group).csv" -notypeinformation
