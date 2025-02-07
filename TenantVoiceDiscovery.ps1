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

function Invoke-UploadTenantDataToBlobStorage {
    [CmdletBinding()]
    param (
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$AccountName,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$ContainerName,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ $_ -ne $null })]
      [PSCustomObject]$TenantData,
  
      [Parameter(Mandatory = $true)]
      [ValidateScript({ -not ([string]::IsNullOrWhiteSpace($_)) })]
      [string]$SasToken
      
    )
  
    begin {
      $output = New-Object System.Collections.Generic.List[PSCustomObject]
      $blobName = $file
  
      $csvData = $TenantData | ConvertTo-Csv -NoTypeInformation
      $csvData = $csvData -replace '"', ''
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvData -join "`n")
  
      $url = "https://saservicemanagerdata.blob.core.windows.net/customerdata/$($blobName)$SasToken"
  
      $headers = @{
        "x-ms-blob-type"         = "BlockBlob"
        "x-ms-blob-content-type" = "application/octet-stream"
      }
    }
  
    process {
      $response = Invoke-WebRequest -Uri $url -Method Put -Headers $headers -Body $bytes -ContentType "application/octet-stream" -ErrorAction SilentlyContinue
  
      if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 300) {
        $output.Add([PSCustomObject]@{
            AccountName   = $AccountName
            ContainerName = $ContainerName
            BlobName      = $blobName
            Uploaded      = $true
            StatusCode    = $response.StatusCode
            Message       = "Successfully uploaded"
          })
      }
      else {
        $output.Add([PSCustomObject]@{
            AccountName   = $AccountName
            ContainerName = $ContainerName
            BlobName      = $blobName
            Uploaded      = $false
            StatusCode    = $response.StatusCode
            Message       = "Failed to upload"
          })
      }
    }
    end {
        return $output
      }
    }

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
if ($CAPPolicies) {
        foreach ($policy in $CAPPolicies) {
            Get-CsOnlineUser | where { $_.TeamsIPPhonePolicy -match $policy } | ForEach-Object {
            $CAPs.Add($_)
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
        CommonAreaPhones = $CAPs.count
        CallQueues    = $callQueueData.Length
        AutoAttendants = $autoAttendantData.Length
      }

      $Export

#Export upload to blob
$file = $verifiedDomain.Substring(0, $verifiedDomain.Length -16) + ".csv"

$token = Read-Host "Please Enter The SAS Token"
$sas = "?sv=2022-11-02&ss=bf&srt=co&sp=rwdlaciytfx&se=2025-05-01T00:23:50Z&st=2023-11-28T17:23:50Z&spr=https&sig=" + $token

Invoke-UploadTenantDataToBlobStorage -TenantData $Export -AccountName "saservicemanagerdata" -ContainerName "customerdata" -SasToken $sas
