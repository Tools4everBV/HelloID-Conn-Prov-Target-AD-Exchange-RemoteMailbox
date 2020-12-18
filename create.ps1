
#Initialize default properties
$c = $configuration | ConvertFrom-Json
$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$success = $False;
$auditMessage = "for person " + $p.DisplayName;

$samaccountname = $p.Accounts.MicrosoftActiveDirectory.sAMAccountName

# Get primary DC
$dc = Get-ADDomainController -Discover -Service PrimaryDC

$user = Get-ADUser -Server $dc.Name -Identity $samaccountname -Properties proxyaddresses,targetaddress,msExchRecipientDisplayType,msExchRecipientTypeDetails,msExchRemoteRecipientType,mail

if([string]::IsNullOrEmpty($user.UserPrincipalName))
{
    Write-Error "No UPN available"
}
else
{
    # Search for targetaddress
    foreach ($proxyaddress in $user.proxyaddresses)
    {
        if($proxyaddress -like "*@<tenant>.mail.onmicrosoft.com")
        {
            $targetaddress = $proxyaddress.Replace("smtp:", "SMTP:")
            break
        }
    }

    if(-Not($dryRun -eq $True)) 
    {
        #Write create logic here
        if([string]::IsNullOrEmpty($user.targetaddress))
        {
            Set-ADUser -Server $dc.Name -Identity $samaccountname -EmailAddress $user.UserPrincipalName -Replace @{msExchRemoteRecipientType="4";msExchRecipientTypeDetails="2147483648";msExchRecipientDisplayType="-2147483642";mailnickname="$samaccountname";targetAddress="$targetaddress"}
        }
        else
        {
            Write-Warning "Not updating AD user because targetaddress is not empty"
        }
    }
    $success = $True
}

$auditMessage = "for person " + $p.DisplayName;

#build up result
$result = [PSCustomObject]@{
	Success= $success;
	AccountReference= $samaccountname;
	AuditDetails=$auditMessage;
};

#send result back
Write-Output $result | ConvertTo-Json -Depth 10