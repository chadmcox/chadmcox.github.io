# Microsoft Office lots and lots of failed logons in Azure AD

I recently spent several weeks troubleshooting this issues. In this case the customer just upgraded the office 365 pro plus desktop apps and in Azure AD we noticed a huge spike in failed authentications in the millions against the Microsoft Office application.  After diving in this is what we found:

Note: Customer also just enabled Hybrid Modern Auth on the on prem exchange servers

* Run the following in Powershell
* Using the AzureADPreview module
```
PS C:\> Get-AzureADApplicationSignInDetailedSummary -filter "AppDisplayName eq 'Microsoft Office'" | where {$_.status.errorcode -eq 500011}

#Will return something like this

Status                  : class SignInStatus {
                            ErrorCode: 500011
                            FailureReason: The resource principal named {name} was not found in the tenant named {tenant}. This can happen if the 
                          application has not been installed by the administrator of the tenant or consented to by any user in the tenant. You might have 
                          sent your authentication request to the wrong tenant.
                            AdditionalDetails: Developer error - the app requested access to a resource (application) that isn't installed in your tenant. 
                          If you expect the app to be installed, you may need to provide administrator permissions to add it. Check with the developers of 
                          the resource and application to understand what the right setup for your tenant is.
                          }

```

The issue is when the Exchange Admins followed the [instructions](https://docs.microsoft.com/en-us/office365/enterprise/configure-exchange-server-for-hybrid-modern-authentication#add-on-premises-web-service-urls-as-spns-in-azure-ad) on configuring modern auth support. There is a step to add all related urls, this must includes cnames and all on prem related urls as well need to be added to the Exchange Online's ServicePrincipalNames
```
$x= Get-MsolServicePrincipal -AppPrincipalId 00000002-0000-0ff1-ce00-000000000000
$x.ServicePrincipalnames.Add("https://mail.corp.contoso.com/")
$x.ServicePrincipalnames.Add("https://owa.contoso.com/")
Set-MSOLServicePrincipal -AppPrincipalId 00000002-0000-0ff1-ce00-000000000000 -ServicePrincipalNames $x.ServicePrincipalNames
```

Once we added all the possible autodiscovery urls, the events almost instantly went away in the event logs and users were able to log into Outlook.

