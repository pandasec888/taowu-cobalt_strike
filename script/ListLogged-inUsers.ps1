<#
.SYNOPSIS
This script will list the logged-in users' RDP Connections History.
Author: 3gstudent@3gstudent
License: BSD 3-Clause
#>
$AllUser = Get-WmiObject -Class Win32_UserAccount
foreach($User in $AllUser)
{
	$RegPath = "Registry::HKEY_USERS\"+$User.SID+"\Software\Microsoft\Terminal Server Client\Servers\"
	Write-Host "User:"$User.Name
	Write-Host "SID:"$User.SID
	Write-Host "Status:"$User.Status
	Try  
    	{ 
		$QueryPath = dir $RegPath -Name -ErrorAction Stop
	}
	Catch
	{
		Write-Host "No RDP Connections History"
		Write-Host "----------------------------------"
		continue
	}
	foreach($Name in $QueryPath)
	{   
		Try  
    		{  
    			$User = (Get-ItemProperty -Path $RegPath$Name -ErrorAction Stop).UsernameHint
    			Write-Host "Server:"$Name
    			Write-Host "User:"$User
    		}
    		Catch  
    		{
			Write-Host "No RDP Connections History"
    		}
	}
	Write-Host "----------------------------------"	
}
