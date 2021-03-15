Function fGet-ExchangeCertificate {
	<#
		.NOTES
			Name: 		fGet-ExchangeCertificate
			Author: 	Roger Buchser
			Version:	2021.03.12
			
		.SYNOPSIS
			Alternative Command for 'Get-ExchangeCertificate'.
			
		.DESCRIPTION
			The Command 'Get-ExchangeCertificate' in Exchange 2016 does not return some data, that was returned with the same
			Command under Exchange 2010. Example: It is not possible to receive the Services information or the Friendly Name.
			This Function will do that with a Trick. Different or better ways to receive that information are Welcome... ;-)
			
			I know, this is a unconventional Method to get the Certificatie Data. But i did not found any other Solution. 
			For me it works...
		
		.PARAMETER Servers
			Select Exchange Servers to get Exchange Certificate Information

		.LINK
			https://github.com/rbuchser/fGet-ExchangeCertificate
		
		.EXAMPLE
			fGet-ExchangeCertificate
			Will return an Certificate Overview from all Mailbox Servers
			
		.EXAMPLE
			fGet-ExchangeCertificate -Servers <MailboxServer16>
			Will return an Certificate Overview from Mailbox Server 16.
			
		.EXAMPLE
			fGet-ExchangeCertificate -Servers <MailboxServer1>,<MailboxServer5>,<MailboxServer14>
			Will return an Certificate Overview from Mailbox Servers 1, 5 and 14.
	#>
	
	PARAM (
		[Parameter(Position=0,Mandatory=$False)][String[]]$Servers
	)
	
	###############################################################################
	# =========================================================================== #
	# Connect to Exchange Server                                                  #
	# =========================================================================== #
	###############################################################################
	# Remove Closed & Broken Exchange PS-Session
	Get-PsSession | Where {($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.State -eq "Closed")} | Remove-PsSession
	Get-PsSession | Where {($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.State -eq "Broken")} | Remove-PsSession
	
	# Create new Exchange PS-Session using HTTPS Method, if no Exchange PS-Session exists
	If (!(Get-PsSession | where {($_.ConfigurationName -eq "Microsoft.Exchange") -AND ($_.State -eq "Opened")})) {
		[String[]]$ExchangeConnectionSequence = @("$($Servers[0])","$($Servers[1])","$($Servers[2])","OUTLOOK","MAIL")
		ForEach ($ExchangeConnection in $ExchangeConnectionSequence) {
			Write-Host "`nTry to to connect to Exchange Server `'https://$ExchangeConnection.$([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name)/PowerShell`'" -f White
			Try {
				$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://$ExchangeConnection.$([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name)/PowerShell/ -Authentication Negotiate -ErrorAction Stop
				Import-PSSession $Session -DisableNameChecking -ErrorAction Stop | Out-Null
				Write-Host "Successfully connected to Exchange Server over `'$($Session.ComputerName)`'" -f DarkGreen
				Break
			} Catch {
				Write-Host "Cannot connect to `'https://$ExchangeConnection.$([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name)/PowerShell`'" -f DarkRed
			}
		}
	}
	
	# If Paramter -Servers is not used, the get Certificate Information from all Mailbox Servers
	If (!($Servers)) {
		$Servers = (Get-MailboxServer | Sort Name).Name
	}
		
	
	###############################################################################
	# =========================================================================== #
	# Effective Code starts here                                                  #
	# =========================================================================== #
	###############################################################################
	# Create Result Variable
	$ExchangeCertificateOverview = @()
	
	# Start first Job on all Servers
	$i = 1
	ForEach ($Server in $Servers) {
		Write-Progress -Activity "Starting Jobs to get Exchange Certificates on total $($Servers.Count) Servers. Please wait..." -Status "Processing $i/$($Servers.Count) Server" -PercentComplete ($i/$Servers.Count*100);$i++
		Try {
			$Job1 = Start-Job -Name "$Server ExCerts"     -Argumentlist $Server -ScriptBlock {Powershell.exe -Noprofile -command "& {Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue; `$FormatEnumerationLimit = -1; Get-ExchangeCertificate -Server $Args -ErrorAction SilentlyContinue | Where {(`$_.Services -match 'IIS') -AND (`$_.IsSelfSigned -eq `$False)} -ErrorAction Stop | Select FriendlyName,HasPrivateKey,IsSelfSigned,Issuer,PublicKeySize,RootCAType,Services,Status,Subject}; Exit"} -ErrorAction Stop
		} Catch {
			$Job1 = Start-Job -Name "$Server ExCerts"     -Argumentlist $Server -ScriptBlock {Get-ExchangeCertificate -Server $Args | Where {($_.Services -match 'IIS') -AND ($_.IsSelfSigned -eq $False)} -ErrorAction SilentlyContinue | Select FriendlyName,HasPrivateKey,Issuer,Subject}
		}
	}
	Write-Progress -Activity "Starting Jobs to get Exchange Certificates on total $($Servers.Count) Servers. Please wait..." -Completed
	
	# Start second Job on all Servers
	$i = 1
	ForEach ($Server in $Servers) {
		Write-Progress -Activity "Starting Jobs to get Exchange Certificate Domains on total $($Servers.Count) Servers. Please wait..." -Status "Processing $i/$($Servers.Count) Server" -PercentComplete ($i/$Servers.Count*100);$i++
		Try {
			$Job2 = Start-Job -Name "$Server CertDomains" -Argumentlist $Server -ScriptBlock {Powershell.exe -Noprofile -command "& {Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue; (Get-ExchangeCertificate -Server $Args -ErrorAction SilentlyContinue | Where {(`$_.Services -match 'IIS') -AND (`$_.IsSelfSigned -eq `$False)} -ErrorAction Stop | Select -ExpandProperty CertificateDomains).Domain -Join ','}; Exit"} -ErrorAction Stop
		} Catch {
			$Job2 = Start-Job -Name "$Server CertDomains" -Argumentlist $Server -ScriptBlock {(Get-ExchangeCertificate -Server $Args | Where {($_.Services -match 'IIS') -AND ($_.IsSelfSigned -eq $False)} -ErrorAction SilentlyContinue | Select -ExpandProperty CertificateDomains).Domain -Join ','}
		}
	}
	Write-Progress -Activity "Starting Jobs to get Exchange Certificate Domains on total $($Servers.Count) Servers. Please wait..." -Completed
	
	# Wait until all Jobs are finished
	$TotalJobs = (Get-Job | Measure).Count
	Do {
		$JobsCompleted = (Get-Job | Where {$_.State -eq "Completed"} | Measure).Count
		Write-Progress -Activity "Waiting for Jobs. Please wait..." -Status "$JobsCompleted/$TotalJobs Jobs completed" -PercentComplete ($JobsCompleted/$TotalJobs*100)
		Start-Sleep 2
	} Until ((Get-Job | Where {$_.State -eq "Running"} | Measure).Count -eq 0)
	Write-Progress -Activity "Waiting for Jobs. Please wait..." -Completed
	
	# Receive Jobs from Server & create a new Object
	$i = 1
	ForEach ($Server in $Servers) {
		Write-Progress -Activity "Receiving Jobs for Exchange Certificates on total $($Servers.Count) Servers. Please wait..." -Status "Processing $i/$($Servers.Count) Server" -PercentComplete ($i/$Servers.Count*100);$i++
		$Obj = New-Object PsObject
		$Obj | Add-Member NoteProperty -Name ServerName -Value $Server
		Try {
			$ExCertInfos = Get-Job -Name "$Server ExCerts" -ErrorAction Stop | Wait-Job | Receive-Job
			$ExCertDomains = Get-Job -Name "$Server CertDomains" -ErrorAction Stop | Wait-Job | Receive-Job
			$Obj | Add-Member NoteProperty -Name FriendlyName       -Value ($ExCertInfos | Select-String FriendlyName).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name CertificateDomains -Value $ExCertDomains
			$Obj | Add-Member NoteProperty -Name HasPrivateKey      -Value ($ExCertInfos | Select-String HasPrivateKey).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name IsSelfSigned       -Value ($ExCertInfos | Select-String IsSelfSigned).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name Issuer             -Value ($ExCertInfos | Select-String Issuer).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name PublicKeySize      -Value ($ExCertInfos | Select-String PublicKeySize).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name RootCAType         -Value ($ExCertInfos | Select-String RootCAType).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name Services           -Value ($ExCertInfos | Select-String Services).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name Status             -Value ($ExCertInfos | Select-String Status).ToString().Split(":")[1].Trim()
			$Obj | Add-Member NoteProperty -Name Subject            -Value ($ExCertInfos | Select-String Subject).ToString().Split(":")[1].Trim()
			$ExchangeCertificateOverview += $Obj
		} Catch {
			$Obj | Add-Member NoteProperty -Name FriendlyName       -Value "N/A"
			$Obj | Add-Member NoteProperty -Name CertificateDomains -Value "N/A"
			$Obj | Add-Member NoteProperty -Name HasPrivateKey      -Value "N/A"
			$Obj | Add-Member NoteProperty -Name IsSelfSigned       -Value "N/A"
			$Obj | Add-Member NoteProperty -Name Issuer             -Value "N/A"
			$Obj | Add-Member NoteProperty -Name PublicKeySize      -Value "N/A"
			$Obj | Add-Member NoteProperty -Name RootCAType         -Value "N/A"
			$Obj | Add-Member NoteProperty -Name Services           -Value "N/A"
			$Obj | Add-Member NoteProperty -Name Status             -Value "N/A"
			$Obj | Add-Member NoteProperty -Name Subject            -Value "N/A"
			$ExchangeCertificateOverview += $Obj
		}
	}
	Write-Progress -Activity "Receiving Jobs for Exchange Certificates on total $($Servers.Count) Servers. Please wait..." -Completed
	
	# Remove existing Jobs
	Get-Job | Remove-Job
	
	# Return Result Object
	Return $ExchangeCertificateOverview
}
