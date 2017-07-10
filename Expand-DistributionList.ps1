Param (
    [Parameter(Mandatory=$true)][string]$DistGroup
)

#Import-Exchange # Import Custom Exchange Management PS Module

Function Get-Type ( $member ) {
  try {
    $recipient = Get-Recipient -ResultSize Unlimited $member.SamAccountName #-ErrorAction SilentlyContinue
    if ( ($recipient).RecipientType -eq "UserMailbox" ) {
      $MyReturn = "UserMailbox"
    } elseif ( ($recipient).RecipientType -eq "MailUniversalSecurityGroup" ) {
      $MyReturn = "MailUniversalSecurityGroup"
    } elseif ( ($recipient).RecipientType -eq "MailUser" ) {
      if ($debug) { Write-Host "Debug :: Migrated to Office 365" }
      $MyReturn = "MailUser"
    } else {
      $MyReturn = "NotAThing"
    }
  } catch {
    #ManagementObjectNotFoundException
    Write-Host "Debug :: Recipient Not Found...which is weird."
  }

  Return $MyReturn
}

Function GetMailboxProperties ( $User ) {
    if ($debug) {Write-Host "Debug :: Processing Mailox for $($User)"}
    $mbox = Get-Mailbox $User #-ErrorAction SilentlyContinue
    $row = @()
    $row = New-Object PSObject; 
    $row | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $mbox.PrimarySMTPAddress
    $row | Add-Member -MemberType NoteProperty -Name "BadItemCount" -Value "0" 
    $row | Add-Member -MemberType NoteProperty -Name "StaffNumber" -Value $mbox.SamAccountName
    $row | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $mbox.DisplayName 
    $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinGB" -Value "0" 
    $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinMB" -Value "0"
    if ($Debug) { Write-Host "Debug :: ROW: $($row.emailAddress)" }
    Return , $row # Returns a row of the user properties 
}

Function CheckDuplicate ($value) {
    For ( $n=0; $n -lt $Global:MasterList.length; $n++ ) {
        if ( $Global:MasterList[$n].EmailAddress -eq $value ) { if ($Verbose) { Write-host "TRUE: $($Global:MasterList[$n].EmailAddress) matches $($value)" }; Return $True } 
    }
    if ($Verbose) { Write-Host "FALSE: Found no duplicate for $value" }
    Return $False 
}

# Expand a dist list into its component members
Function EnumMember ( $distListMember ) {
    if ($debug) {Write-Host "Debug :: Processing Member $($distListMember.SamAccountName)"}
    if ( (Get-Type $distListMember) -eq "MailUser" ) {
        if ( ! ( CheckDuplicate $distListMember.PrimarySMTPAddress ) ) {
            if ($debug) { Write-Host "Debug :: MIGRATED or DELETED"}
            $row = @()
            $row = New-Object PSObject;
            $row | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value "$($distListMember.PrimarySMTPAddress)"
            $row | Add-Member -MemberType NoteProperty -Name "BadItemCount" -Value "ALREADY MIGRATED" 
            $row | Add-Member -MemberType NoteProperty -Name "StaffNumber" -Value "-"
            $row | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "-"
            $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinGB" -Value "-"
            $row | Add-Member -MemberType NoteProperty -Name "TotalItemSizeinMB" -Value "-"
            if ($Verbose) { Write-host "$row" }
            $Global:MasterList  += $row
            $row = $null
          }
      }
      if ( (Get-Type $distListMember) -eq "UserMailbox" ) {
        if ( ! ( CheckDuplicate $distListMember.PrimarySMTPAddress ) ) {
            if ($Debug) { Write-Host "Debug :: User Mailbox: $($distListMember.DisplayName)" }
            $row = @()
            $row = GetMailboxProperties $distListMember.SamAccountName
            if ($Verbose) { Write-host "$row" }
            $Global:MasterList += $row        
        }
    }

    if ( (Get-Type $distListMember) -eq "MailUniversalSecurityGroup" ) {
        $list = Get-DistributionGroupMember $distListMember.SamAccountName #-ErrorAction SilentlyContinue
        Foreach ( $listMember in $list) {
            if ( $verbose ) { Write-Host "Processing Nested Dist List $($listMember)" }
            EnumMember $listMember
            Continue
        }
    } 
    Return
}

$debug             = $False
$Verbose           = $False
$global:MasterList = @()
$distGroupMembers  = Get-DistributionGroupMember $DistGroup

Foreach ($member in $distGroupMembers) {
    EnumMember $member
}
Write-Host "Total recursive unique members: $($MasterList.length)"
$MasterList | Export-CSV -Path ".\$($DistGroup).csv" -notypeinformation
 
