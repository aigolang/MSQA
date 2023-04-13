# @Title: Test EWS OAuth with Application Type Permission (Exchange Online)
# @date: 2020.10.27

########### >>> Request Access_Token >>> ############ 

# EWS OAuth App-Only Authentication

## >>> Customize >>>
param(
    [string] $TargetMailboxUPN = "<UPN>",
    [switch] $TenantID = "",
    [switch] $AppId = "",
    [switch] $AppSecret = "",
    [int]$TopNum = 5
)

Add-Type -Path 'C:\Code\PowerShell\DLLfolder\Microsoft.IdentityModel.Clients.ActiveDirectory.dll';
Add-Type -Path "C:\Code\PowerShell\DLLfolder\Microsoft.Exchange.WebServices.dll";
## <<< Customize <<<

$authString = "https://login.microsoftonline.com/$TenantID";

# this part uses the classes to obtain the necessary security token for performing our operations against the Graph API
$creds = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential" -ArgumentList $AppId, $AppSecret;
$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext"-ArgumentList $authString;
$context = $authContext.AcquireTokenAsync("https://outlook.office365.com/", $creds).Result;
$acc_token = $context.AccessToken;

############ <<< Request Access_Token <<< ###########

# EWS Trace Log >>>
function TraceHandler(){
$sourceCode = @"
    public class ewsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
    {
        public System.String LogFile {get;set;}
        public void Trace(System.String traceType, System.String traceMessage)
        {
            System.IO.File.AppendAllText(this.LogFile, traceMessage);
        }
    }
"@    

   Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $Script:EWSDLL
   $TraceListener = New-Object ewsTraceListener

   return $TraceListener
}

# EWS log Enabled
$EWSService.TraceEnabled = $True
$TraceHandlerObj = TraceHandler
$TraceHandlerObj.LogFile = "C:\temp\EwsLog_$($TargetMailboxUPN).log"
$EWSService.TraceListener = $TraceHandlerObj 
# EWS Trace Log <<<

## Create the Exchange Service object with Oauth creds
$EWSService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2010_SP2
$EWSService.Url= new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx")
#$EWSService.TraceEnabled = $true
$EWSService.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$TargetMailboxUpn)
$EWSService.HttpHeaders.Add("X-AnchorMailbox", $TargetMailboxUpn)

# >>> !!! Use the Access_Token for EWS Object. !!!
$EWSService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($acc_token);
# <<< !!! Use the Access_Token for EWS Object. !!!


$folderId= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $TargetMailboxUPN);
$targetFolder=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWSService,$folderId);

$ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$Results = @()
$fiItems = $null  
do{  
    $fiItems = $targetFolder.FindItems($ItemView)
    # $fiItems = $targetFolder.FindItems($FolderView)  
    foreach($Item in $fiItems.Items){  
        # "RecivedDate : " + $Item.DateTimeReceived   
        # "Subject     : " + $Item.Subject   
        # "Size        : " + $Item.Size

        if ($Results.Count -lt $topNum){

            $myObject1 = [PSCustomObject]@{
                UPN = $targetMailboxUPN
                From = $Item.From
                #To = $Item.To
                RecivedDate = $Item.DateTimeReceived
                Subject = $Item.Subject
            }
        
            $Results = $Results + $myObject1
        }
        else{
            break
        }
            
    }  
    $ItemView.Offset += $fiItems.Items.Count
} while($fiItems.MoreAvailable -eq $true)


Write-Host "------------------ Result Show Top $TopNum ------------------"
$Results | FT

