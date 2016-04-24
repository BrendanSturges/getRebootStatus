<#

Script created by Brendan Sturges, reach out if you have any issues.
This script queries a file the user chooses and checks all servers within to see if the server is in a pending reboot status and exports this to a CSV

#>


 Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension

}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName

$i = 0
foreach($server in $serverList) {
	Try {
		$getRebootStatus = [wmiclass]"\\$server\root\ccm\clientsdk:ccm_clientutilities"
		$result = $getRebootStatus.DetermineifRebootPending() | Select RebootPending
	
		if ($result.rebootpending) {
			#$text = "$server, requires reboot"
			#$text | out-file $filename -append
			
			$props = [ordered]@{
			'Server' = $server
			'Reboot Pending' = $result.rebootpending	
			'Details' = ''
			}
		
			$obj = New-Object -TypeName PSObject -Property $props
		
			} 
	
		else { 
			#$text = "$server, ok"
			#$text | export-csv $filename -append
		 
			$props = [ordered]@{
			'Server' = $server
			'Reboot Pending' = $result.rebootpending
			'Details' = ''			
			}	
		
			$obj = New-Object -TypeName PSObject -Property $props
			}
		}
	Catch {
		if(Test-Connection -ComputerName $server -Count 2 -Quiet)
			{
			$ErrorMessage = $_.Exception.Message
			}
		else
			{
			$ErrorMessage = 'Server is Offline'
			}
		
		$props = [ordered]@{
			'Server' = $server
			'Reboot Pending' = 'ERROR'
			'Details' = $ErrorMessage
			}
			
		$obj = New-Object -TypeName PSObject -Property $props
	
	}
	Finally {
		$data = @()
		$data += $obj
		$data | Export-Csv $fileName -noTypeInformation -append	
	}
	$i++
	Write-Progress -activity "Checking server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)	
}


