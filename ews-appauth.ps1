$TenantId = ""
$AppClientId=""
$pass = ConvertTo-SecureString -AsPlainText ""

$MsalParams = @{
            ClientId = $AppClientId
            TenantId = $TenantId
            ClientSecret = $pass
            Scopes   = "https://outlook.office.com/.default"
}

##################################################################################################################
##################################################################################################################
##################################################################################################################

Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

#Provide the mailbox id (email address) to connect
$MailboxName =""
 
#Get Access Token with the scope "https://outlook.office.com/EWS.AccessAsUser.All"
$EWSAccessToken=(Get-MsalToken @MsalParams).AccessToken

$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
 
#Use Modern Authentication
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken

#Check EWS connection
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$Service.AutodiscoverUrl($MailboxName,{$true})
$Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(2, $MailboxName)
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)

$findResults = $inbox.FindItems($view)

foreach ($item in $findResults.Items)
{
    $item.Load()
    Write-Host $item.Subject
}
