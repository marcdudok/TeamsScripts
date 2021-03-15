<#
NOTES
This script is intended to quickly provision Microsoft Teams (MTR or Surface Hub) resource accounts and assign licenses from a CSV File.

An example CSV file is defined below.
------------
UPN,Alias,DisplayName,AllowExternalInvites,passwordNeverExpires,MTRP
Room1@contoso.com,Room1,Demo Room 1,TRUE,TRUE,FALSE
Room2@contoso.com,Room2,Demo Room 2,TRUE,TRUE,FALSE
--------
"AllowExternalInvites" will allow external domains to book this resource account
"MTRP" defines whether you would like to assign a Microsoft Teams Room Premium license to this account
--------

Official Guidance and scripts with additional error checking and reporting. 
https://docs.microsoft.com/en-us/surface-hub/surface-hub-2s-account#create-account-via-powershell
https://docs.microsoft.com/en-us/surface-hub/appendix-a-powershell-scripts-for-surface-hub
#>

$Error.Clear()
$ErrorActionPreference = "Stop"
$status = @{}
#Get the CSV Input File
$roomCSV = import-csv c:\temp\rooms.csv

try {
    Import-Module ExchangeOnlineManagement
    Import-Module MSOnline
    Import-Module AzureAD
}
catch
{
    write-host "Some dependencies are missing"
    write-host "Please install the Windows PowerShell Module for ExchangeOnlineManagement." -ForegroundColor Red
    write-host "Please install the Windows PowerShell Module for MSOnline" -ForegroundColor Red
    write-host "Please install the Azure Active Directory module for PowerShell" -ForegroundColor Red
    exit 1
}

#Establish connections - User will be prompted for UPN and PW
try {
    Write-Progress "Connecting to ExO"
    #Connect-ExchangeOnline -ShowProgress $true -ShowBanner:$false
    Write-Progress "Connecting to MSOL"
    #Connect-MsolService
    Write-Progress "Connecting to AAD"
    #Connect-AzureAD
}
catch{
    write-host "Connection to one or more services failed. Please validate your credentials." -ForegroundColor Red
    $error[0]
    #Cleanup so we don't get throttled
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

function New-O365ResourceAccount ($UPN, $Alias, $Name, [bool]$AllowExternalInvites, [bool]$passwordNeverExpires, [bool]$MTRP)
{
    #Check if it exists already
    if(!(Get-Mailbox $UPN -ErrorAction SilentlyContinue))
        {
            #Generate a password
            $genPassword = [System.Web.Security.Membership]::GeneratePassword(16, 0)
            #Create the Mailbox
            New-Mailbox -MicrosoftOnlineServicesID $UPN -Alias $Alias -Name $Name -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword (ConvertTo-SecureString -String $genPassword -AsPlainText -Force) | Out-Null
         }
    else
        {
        #Mailbox already Exists but we can continue to configure properties as defined
    }

    #Set Calendar Response and processing options. 
    #There is a delay between mailbox creation and the ability to set Processing Rules.

    do{
        write-progress -Activity "Waiting for Mailbox $($UPN) creation to finish"
        Start-Sleep -seconds 1
        }
    until ((!(Get-CalendarProcessing -Identity $($UPN) -errorAction SilentlyContinue)) -eq $false)

    Set-CalendarProcessing -Identity $($UPN) -AutomateProcessing AutoAccept -AddOrganizerToSubject $false –AllowConflicts $false –DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false
    
    if($AllowExternalInvites){
        Set-CalendarProcessing -Identity $UPN -ProcessExternalMeetingMessages $true
    }

    if($passwordNeverExpires){
        Set-MsolUser -UserPrincipalName $UPN -PasswordNeverExpires $True
    }

    #
    $MTR_SKU = "6070a4c8-34c6-4937-8dfb-39bbc6397a60"
    $MTRP_SKU = "4fb214cb-a430-4a91-9c91-4976763aa78f"

    if($MRTP)
        {
        $licenseTypeSKU = $MTRP_SKU
        }
        else
        {
        $licenseTypeSKU = $MTR_SKU
        }

    #Assign the account a license
    $errorVar = $null
    Set-AzureADUser -ObjectId $UPN -UsageLocation AU
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense 
    $License.SkuId = $licenseTypeSKU
    $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses 
    $Licenses.AddLicenses = $License
    #try{
    #    Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $Licenses -ErrorVariable errorVar
    #    }
    #catch{
    #    write-host "Could not assign a license to the account. Ensure additional licenses are available of specified type prior to executing this script" -ForegroundColor Red
    #    exit 1 
    #} 

    $returnObj = New-Object -TypeName PSObject
    $returnObj | Add-Member -MemberType NoteProperty "UPN" -Value $UPN
    $returnObj | Add-Member -MemberType NoteProperty "Display Name" -Value $Name
    if($genPassword){$returnObj | Add-Member -MemberType NoteProperty "Password" -Value $($genPassword)}else{$returnObj | Add-Member -MemberType NoteProperty "Password" -Value "Pre-existing account"}

    return $returnObj
}

$outObjs = @()
foreach($room in $roomCSV)
{
    #Change input values to BOOL
    if($room.AllowExternalInvites -match "TRUE"){$allowExternalInvites = $true}else{$allowExternalInvites = $false}
    if($room.passwordNeverExpires -match "TRUE"){$passwordNeverExpires = $true}else{$passwordNeverExpires = $false}
    if($room.MTRP -match "TRUE"){$MTRP = $true}else{$MTRP = $false}
    #Call the worker function
    $obj = New-O365ResourceAccount -UPN $room.UPN -Alias $room.alias -Name $room.DisplayName -AllowExternalInvites $allowExternalInvites -passwordNeverExpires $passwordNeverExpires -MTRP $MTRP
    $outObjs += $obj
}
$outObjs | Format-Table