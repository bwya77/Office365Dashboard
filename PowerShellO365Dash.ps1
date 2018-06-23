Function New-GraphToken
{
	#Requires -Module AzureRM
	[CmdletBinding()]
	Param (
		$TenantName = 'bwya77.onmicrosoft.com'
	)
	
	try
	{
		Import-Module AzureRM -ErrorAction Stop
	}
	catch
	{
		Write-Error 'Can''t load AzureRM module.'
		break
	}
	
	$clientId = "1950a258-227b-4e31-a9cf-717495945fc2" #PowerShell ClientID
	$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
	$resourceAppIdURI = "https://graph.windows.net"
	$authority = "https://login.windows.net/$TenantName"
	$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
	#$authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$redirectUri, "Auto")
	$authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId, $redirectUri, "Always")
	
	@{
		'Content-Type'	      = 'application\json'
		'Authorization'	      = $authResult.CreateAuthorizationHeader()
	}
}

$TenantName = "bwya77.onmicrosoft.com"
$GraphToken = New-GraphToken -TenantName $TenantName

$Colors = @{
	BackgroundColor   = "#FF252525"
	FontColor		  = "#FFFFFFFF"
}

$NavBarLinks = @((New-UDLink -Text "<i class='material-icons' style='display:inline;padding-right:5px'>favorite_border</i> PowerShell Pro Tools" -Url "https://poshtools.com/buy-powershell-pro-tools/"),
	(New-UDLink -Text "<i class='material-icons' style='display:inline;padding-right:5px'>description</i> Documentation" -Url "https://adamdriscoll.gitbooks.io/powershell-tools-documentation/content/powershell-pro-tools-documentation/about-universal-dashboard.html"))


Start-UDDashboard -Wait -Port 8081 -Content {
	New-UDDashboard -NavbarLinks $NavBarLinks -Title "Office 365 Dashboard" -NavBarColor '#FF1c1c1c' -NavBarFontColor "#FF55b3ff" -BackgroundColor "#FF333333" -FontColor "#FFFFFFF" -Content {
		New-UDRow{
			New-UDColumn -Size 4 {
				New-UDMonitor -Title "Total Users" -Type Line -DataPointHistory 20 -RefreshInterval 15 -ChartBackgroundColor '#5955FF90' -ChartBorderColor '#FF55FF90' @Colors -Endpoint {
					(Invoke-RestMethod -Uri "https://graph.windows.net/$TenantName/users/?api-version=1.6" -Headers $GraphToken -Method Get | Select-Object -ExpandProperty Value).Count | Out-UDMonitorData
				}
			}
			New-UDColumn -Size 4 {
				New-UDMonitor -Title "Total Groups" -Type Line -DataPointHistory 20 -RefreshInterval 15 -ChartBackgroundColor '#5955FF90' -ChartBorderColor '#FF55FF90' @Colors -Endpoint {
					(Invoke-RestMethod -Uri "https://graph.windows.net/$TenantName/groups/?api-version=1.6" -Headers $GraphToken -Method Get | Select-Object -ExpandProperty Value).Count | Out-UDMonitorData
				}
			}
			New-UDColumn -Size 4 {
				New-UDGrid -Title "Users Forced to Change Password at Next Login" @Colors -Headers @("User") -Properties @("User") -AutoRefresh -RefreshInterval 20 -Endpoint {
					$PWUsers = Invoke-RestMethod -Uri "https://graph.windows.net/$TenantName/users/?api-version=1.6" -Headers $GraphToken -Method Get | Select-Object -ExpandProperty Value | Where-Object { $_.passwordProfile -like "*forceChangePasswordNextLogin=True*" }
					$UserData = @();
					foreach ($PWUser in $PWUsers)
					{
						$UserData += [PSCustomObject]@{ "User" = ($PWUser).displayName }
					}
					$UserData | Out-UDGridData
				}
			}
		}
		New-UDRow{
			New-UDColumn -Size 7{
				New-UdChart -Title "Licenses" -Type Bar -AutoRefresh -RefreshInterval 7 @Colors -Endpoint {
					$Licenses = Invoke-RestMethod -Uri "https://graph.windows.net/$TenantName/subscribedSkus/?api-version=1.6" -Headers $GraphToken -Method Get | Select-Object -ExpandProperty Value | Select-Object SkuPartNumber, ConsumedUnits -ExpandProperty PrepaidUnits | Where-Object { $_.enabled -lt 10000 }
					$LicenseData = @();
					foreach ($License in $Licenses)
					{
						$Overage = (($License).enabled) - (($License).consumedUnits)
						$LicenseData += [PSCustomObject]@{ "License" = ($License).skuPartNumber; "ConsumedUnits" = ($License).consumedUnits; "EnabledUnits" = ($License).enabled; "UnUsed" = $Overage }
					}
					
					$LicenseData | Out-UDChartData -LabelProperty "License" -Dataset @(
						New-UdChartDataset -DataProperty "ConsumedUnits" -Label "Assigned Licenses" -BackgroundColor "#80962F23" -HoverBackgroundColor "#80962F23"
						New-UdChartDataset -DataProperty "EnabledUnits" -Label "Total Licenses" -BackgroundColor "#8014558C" -HoverBackgroundColor "#8014558C"
						New-UDChartDataset -DataProperty "UnUsed" -Label "Un-Used Licenses" -BackgroundColor "#803AE8CE" -HoverBackgroundColor "#803AE8CE"
					)
				}
			}
			New-UDColumn -Size 4{
				New-UDGrid -Title "Domains" @Colors -Headers @("Domains") -Properties @("Domains") -AutoRefresh -RefreshInterval 20 -Endpoint {
					$Domains = Invoke-RestMethod -Uri "https://graph.windows.net/$TenantName/domains/?api-version=1.6" -Headers $GraphToken -Method Get | Select-Object -ExpandProperty Value | Where-Object { $_.name -notlike "*onmicrosoft.com*" }
					$Domaindata = @();
					foreach ($Domain in $Domains)
					{
						$DomainData += [PSCustomObject]@{ "Domains" = ($Domain).name }
					}
					$DomainData | Out-UDGridData
				}
			}
		}
	}
}
