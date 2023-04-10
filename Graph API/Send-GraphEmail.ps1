<###############################################################################################################
 Script Name: Send-GraphEmail.ps1
 Created By : Gurdeep Gurdev Singh
 Parameters : None
 Description: This Script will send Email form O365 Mailbox using Graph API
 Warning: Test the Script before Running in your environment.
 Version:
   0.1 - Base Script
Important Note:
1. It is recommended to not hard code Client ID and Tenant ID and Secret in the script, instead use, 
   Azure Key Vault/PowerShell Value or encrypted JML and Import it in the script.
2. Script will need a App Registered in Azure AD with Mail.Send Application Permissions. This service
   principal will be able to send it as any mailbox in the tenant. To restrict the mailbox to specific mailbox
   or mailboxes. Create Application access policy in Exchange Online and Scope it with Security Group with 
   mailboxes as members of the group.

################################################################################################################>

$Global:ClientID = "Enter" #$ClientID
$Global:TenantID = "Enter" #$TenantID
$Global:secret = "Enter your secret here" # Secret
Function GetToken ### Function to generate token
{
  try{
    $ReqTokenBody = @{
      Grant_Type = "client_credentials"
      Scope = "https://graph.microsoft.com/.default"
      client_Id = $ClientID
      Client_Secret = $Secret
    }
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($TenantID)/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
    return $TokenResponse
  }catch {return $null}
}#End of Function GetToken
$Global:Token = GetToken
if($Global:Token){
  $Headers = @{
    'Content-Type'  = "application\json"
    'Authorization' = "Bearer $($Token.access_token)" 
  }
  $TO = "TestUser@Test.com"
  $From = "Gurdeep@Test.com"
  $Subject = "Test Email from Graph API"   
  $htmlbody="<html>
             <style>
             BODY{font-family: Arial; font-size: 10pt;}
	         H1{font-size: 22px;}
	         H2{font-size: 18px; padding-top: 10px;}
	         H3{font-size: 16px; padding-top: 8px;}
             </style>
             <body>
             <h1>This email was sent from Graph API.</h1><br> 
             <p><strong>Generated:</strong> $(Get-Date -Format g)</p>  
             </body>"
  $MessageParams = @{
          "URI"         = "https://graph.microsoft.com/v1.0/users/{0}/sendMail" -f $From
          "Headers"     = $Headers
          "Method"      = "POST"
          "ContentType" = 'application/json'
          "Body" = (@{
                "message" = @{
                "subject" = $Subject
                "body"    = @{
                    "contentType" = 'HTML' 
                     "content"     = $htmlbody }

           "toRecipients" = @(
           @{
             "emailAddress" = @{"address" = $To }
           } ) 
         }
      }) | ConvertTo-JSON -Depth 6
   }
   # Send the message
   Invoke-RestMethod @Messageparams
}else{
  Write-Host -ForegroundColor Red -BackgroundColor Cyan "Error: Failed to get token."
}
