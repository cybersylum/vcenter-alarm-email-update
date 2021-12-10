<#
This script will replace any and all e-mail addresses configured on alarms with a single new address

Modify the 2 values below to fit the needs in your environment
#>

$vCenter="vcsa.cybersylum.com"
$NewEmailAddress = "arron@cybersylum.com"


$null=Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

write-host 
write-host "This script will replace any and all existingt e-mail addresses used as recipients in vCenter Server Alarms with"
write-host "a single new e-mail address.  Edit this script to set the new e-mail address ($NewEmailAddress)."
write-host 

#initialize to empty - will prompt for value later
$AlarmTargetName = ""

write-host "Logging into vCenter Server - $vCenter"
$vc=Connect-VIServer -server $vCenter 
if (!$vc) {
    write-host -ForegroundColor 'red' "Unable to connect to vCetnter.  Please try again"
    Write-Host
    exit
}

$AlarmTargetName = Read-Host -Prompt 'Enter the name of the vCenter Alarm you wish to change E-mail Recipients on [leave blank and hit ENTER to change all]'

if ($AlarmTargetName) {
    $VCAlarms=Get-AlarmDefinition $AlarmTargetName  
    $alarmcount = $VCAlarms.count 
    if (!$VCAlarms) {
        write-host "Please try again"
        exit
    } else { 
        write-host "Updating $alarmcount alarm(s) - $AlarmTargetName"
    }
} else {
    $VCAlarms=Get-AlarmDefinition 
    $alarmcount = $VCAlarms.count
    if (!$VCAlarms) {
        write-host "Please try again"
        exit
    } else {  
        write-host "Updating all defined alarms - $alarmcount found"
    }
}

Write-Host
$proceed = Read-Host "Do you wish to proceed ? (Y to proceed, hit enter to cancel)"

if ($proceed -ne "Y") { 
    write-host "Script exiting - no changes made"
    exit
 }

 write-host "will update now..."

foreach ($alarm in $VCAlarms) {
    #write-host "Checking $alarm.name"
    $alarmaction=Get-AlarmAction $alarm
    if ($alarmaction.actiontype -eq "SendEmail") {
        write-host "$alarm is configured to send e-mail.  Updating..."
        $AlarmAction = Get-AlarmAction -AlarmDefinition $alarm
        $mail = $AlarmAction | where {$_.ActionType -eq 'SendEmail'}
        if ($mail.Subject) {
            $null=Remove-AlarmAction -AlarmAction $mail -Confirm:$false
            $null=New-AlarmAction -AlarmDefinition $alarm -Email -To $NewEmailAddress -Subject $mail.Subject -Confirm:$false
        } else {
            write-host -ForegroundColor "orange"  "Please check $alarm.name manually.  The configured e-mail subject may be blank"
        }
    }

}

disconnect-viserver $vc -confirm:$false