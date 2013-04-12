<#  -------------------------------------------------
Name: 			WAURemove
Description: 	Windows Automatic Update Remover
				Removes specific installed KBs from
                Domain

Author:			Ben0xA
Shout Outs:		mwjcomputing
Params:
	$kbs			Comma separated values of KB numbers
	$outputFile		The output file to save results
    $psdir          Path to psexec.exe
	$computer		Specifies a single computer to scan
-----------------------------------------------------
#>

Param(
	[Parameter(Mandatory=$true,Position=1)]
	[string]$kbs,
	
	[Parameter(Mandatory=$false,Position=2)]
	[string]$outputFile,
    
    [Parameter(Mandatory=$false,Position=3)]
	[string]$psdir,
	
	[Parameter(Mandatory=$false,Position=4)]
	[string]$computer
)

Function Get-Pcs{
	$domain = New-Object System.DirectoryServices.DirectoryEntry
	
	$ds = New-Object System.DirectoryServices.DirectorySearcher
	$ds.SearchRoot = $domain
	$ds.Filter = ("(objectCategory=computer)")
	$ds.PropertiesToLoad.Add("name")
	
	$rslts = $ds.FindAll()
	return $rslts
}

Function Remove-KBs($pcname, $kb){
    $wusacmd = $psdir + ' -d -s \\' + $pcname + ' cmd /c "wusa.exe /uninstall /kb:' + $kb + ' /quiet /norestart"'
    Write-Host ("Sending removal command " + $wusacmd)
    Invoke-Expression $wusacmd
}

Function Get-KBs($pcname){
	$rslt = ""
	$qfe = Get-WmiObject -Class Win32_QuickFixEngineering -Computer $pcname -ErrorVariable myerror -ErrorAction SilentlyContinue
	if($myerror.count -eq 0) {
		foreach($kb in $kbItems){
			$installed = $false
			$kbentry = $qfe | Select-String $kb
			if($kbentry){
                Write-Host("KB $kb found. Attempting to uninstall, please wait...")
				$rmrslt = Remove-KBs $pcname $kb
                $rslt += "$pcname,$kb,Uninstall Sent`r`n"            
			}
            else {
                $rslt += "$pcname,$kb,Not Installed`r`n"
            }
		}
	}
	else{
		$rslt += "$pcname,$kb,RPC_Error`r`n"
	}
	return $rslt
}

# Begin Program Flow

Clear-Host
Write-Host "WAURemove"
Write-Host "Written By: @Ben0xA"
Write-Host "Huge thanks to @mwjcomputing!`r`n"
Write-Host "Looking for KBs $kbs"
if(-not $outputFile){
	Write-Host "Sending output to the screen. Use -outputFile name to save to a file.`r`n"
}
else {
	Write-Host "Will save csv results to $outputFile. Query messages will only appear on the screen.`r`n"
}

$wumaster = "PC Name,KB,Status`r`n"
$kbItems = $kbs.Split(",")
if(-not $computer){
	$pcs = Get-PCs
	
	foreach($pc in $pcs){
		$pcname = $pc.Properties.name
		
		if($pcname){
			Write-Host "Querying $pcname, please wait..."
			$wumaster += Get-KBs($pcname)
		}	
	}
}
else{
	Write-Host "Querying $computer, please wait..."
	$wumaster += Get-KBs($computer)
}

if(-not $outputFile){
	Clear-Host
	$wumaster
}
else {
	$wumaster| Out-File $outputFile
	Write-Host "Output saved to $outputFile"
}

#End Program