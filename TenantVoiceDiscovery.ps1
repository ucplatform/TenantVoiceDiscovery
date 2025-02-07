<#
.SYNOPSIS
    Pulls all voice related tenant counts for a customers enviroment

.DESCRIPTION
    Pulls all voice related tenant counts for a customers enviroment, it requires the MicrosoftTeams module to be installed and connected to the tenant. You also need to have the correct permissions to run this script. 

.INPUTS
    Requires the SAS token to be entered when the script is run

.OUTPUTS
    Output to HALO blob storage

.NOTES
    Author:  Adam Smith
    Email: adam.smith@sipcom.com
	Version: 1.1
#>

Install-Module -Name MicrosoftTeams -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
Connect-MicrosoftTeams

#Customer Domain
[string]$verifiedDomain = (Get-CsTenant | Select -ExpandProperty VerifiedDomains | Where { $_.Name -like "*.onmicrosoft.com" -and $_.Name -notlike "*.mail.onmicrosoft.com" }).Name

#User Counts
$Users = Get-CsOnlineUser | Select-Object AccountEnabled, EnterpriseVoiceEnabled, OnlineVoiceRoutingPolicy, LineUri,identity, DisplayName, TeamsIPPhonePolicy
$ActiveTenantUsers = $Users | Where-Object {$_.AccountEnabled -eq $true}
$EntVoiceTenantUsers = $Users | Where-Object {$_.EnterpriseVoiceEnabled -eq $true}
$VRPTenantUsers = $Users | Where-Object {$_.OnlineVoiceRoutingPolicy -ne $null}
$LineUsers = $Users | Where-Object {$_.LineUri -ne $null}

# Global Voice Routing Policy check
$GlobalVRP = "0"
$voiceRoutingPolicyGlobal = Get-CsOnlineVoiceRoutingPolicy -Identity "Global"
          if ($voiceRoutingPolicyGlobal.OnlinePstnUsages) {
                $GlobalVRP = "1"
          }

#Call Queues and Auto Attendants
$callQueueData = Get-CsCallQueue -WarningAction SilentlyContinue
$autoAttendantData = Get-CsAutoAttendant -WarningAction SilentlyContinue
$CAPPolicies = (Get-CsTeamsIPPhonePolicy | where { $_.signinmode -eq "CommonAreaPhoneSignIn" }).identity.replace('Tag:','')

# count CAPs
$CAPs = 0
if ($CAPPolicies) {
        foreach ($policy in $CAPPolicies) {
            $Users | where { $_.TeamsIPPhonePolicy -match $policy } | ForEach-Object {
            $CAPs ++
          }
        }
}
# Calling plan numbers and assigned
$activeLicensedUsersWithCallingPlans = (Get-CsPhoneNumberAssignment -NumberType CallingPlan -ActivationState Activated | where { $_.AssignedPstnTargetId -ne $null }).count
$activeCallingPlanDDIs = (Get-CsPhoneNumberAssignment -NumberType CallingPlan -ActivationState Activated).count

# Operator Connect numbers and assigned
$activeLicensedUsersWithOperatorConnect = (Get-CsPhoneNumberAssignment -NumberType OperatorConnect -ActivationState Activated | where { $_.AssignedPstnTargetId -ne $null }).count
$activeOperatorConnectDDIs = (Get-CsPhoneNumberAssignment -NumberType OperatorConnect -ActivationState Activated).count

#Total Output
$Export = @{
        CustomerTenant          = $verifiedDomain
        TenantActiveUsers   = $ActiveTenantUsers.Count
        GlobalVRPExists = $GlobalVRP
        EnterpriseVoiceEnabledUsers = $EntVoiceTenantUsers.Count
        UsersAssignedVRP = $VRPTenantUsers.Count
        UsersWithLineURI = $LineUsers.Count
        CallingPlanNumbers = $activeCallingPlanDDIs
        CallingPlanAssigned = $activeLicensedUsersWithCallingPlans
        OperatorConnectNumbers = $activeOperatorConnectDDIs
        OperatorConnectAssigned = $activeLicensedUsersWithOperatorConnect
        CommonAreaPhones = $CAPs
        CallQueues    = $callQueueData.Length
        AutoAttendants = $autoAttendantData.Length
      }

foreach ($key in $Export.Keys) {
    Write-Output "$key $($Export[$key])"
}

#Export upload to blob
$file = $verifiedDomain.Substring(0, $verifiedDomain.Length -16) + ".csv"

$token = Read-Host "Please Enter The SAS Token"
$sas = "?sv=2022-11-02&ss=bf&srt=co&sp=rwdlaciytfx&se=2025-05-01T00:23:50Z&st=2023-11-28T17:23:50Z&spr=https&sig=" + $token

$csvData = $Export.GetEnumerator() | select Key, Value | ConvertTo-Csv -NoTypeInformation
      $csvData = $csvData -replace '"', ''
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvData -join "`n")
  
      $url = "https://saservicemanagerdata.blob.core.windows.net/customerdata/$($file)$sas"
  
      $headers = @{
        "x-ms-blob-type"         = "BlockBlob"
        "x-ms-blob-content-type" = "application/octet-stream"
      }

      $response = Invoke-WebRequest -Uri $url -Method Put -Headers $headers -Body $bytes -ContentType "application/octet-stream" -ErrorAction SilentlyContinue

    if($response.StatusCode -eq 201) {
        Write-Host "File uploaded successfully"
    } else {
        Write-Host "File upload failed"
    }

