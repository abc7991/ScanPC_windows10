#donwload RSAT from Microsoft
function installRSAT {
    $parameter = @{
        Uri = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS_1803-x64.msu'
        OutFile = "C:\WindowsTH-RSAT_WS_1803-x64.msu"
    }
    Invoke-WebRequest @parameter
    Test-Path -Path "C:\WindowsTH-RSAT_WS_1803-x64.msu"
    #install RSAT
    wusa.exe /quite /noreboot C:\WindowsTH-RSAT_WS_1803-x64.msu
}
#check msc
$dsaPath = "C:\WINDOWS\system32\dsa.msc"
$checkdsaPath = Test-Path -Path $dsaPath;
if ("$checkdsaPath" -eq "False"){
    installRSAT;
    Write-Warning "If the RSAT cannot download and install automatically, please do it manual."
}
#define current folder

#credential
function SetPasswordAutoSave{
    param([parameter(Mandatory=$true,Position=1)]$username)
    $SecuritypathPassword = "C:\Windows\Temp\pwds.txt"
    if (Test-Path "$securitypathPassword"){
        $password = Get-Content "$SecurityPathPassword" | ConvertTo-SecureString
        $username = $username
    }else{
        (get-credential $username).password | ConvertFrom-SecureString | set-content "$SecurityPath"
        $password = Get-Content "$SecurityPath" | ConvertTo-SecureString
    }
    $script:credential = New-Object System.Management.Automation.PsCredential($username,$password)
    #check credential
    $error.clear()
    New-PSDrive -Name "K" -PSProvider FileSystem -Root \\$env:COMPUTERNAME\C$ -Credential $credential -Persist -ErrorAction SilentlyContinue
    Remove-PSDrive -Force -Name "K" -ErrorAction SilentlyContinue
    if($error[0] -ne $null){
        Write-Warning "Error! Username or password is invalid. Try again"
        Remove-Item -Path $SecuritypathPassword -Force
        exit;
    }else{
        Clear-Host | Start-Sleep 1
        Write-Host "Authenticated!" -ForegroundColor Green;
    }
}
function CreateLogFolder{
    param ($LogFolder,$getdate,$getmonth)
    $ou = (Get-ADOrganizationalUnit -Identity $(($adComputer = Get-ADComputer -Identity $env:COMPUTERNAME).DistinguishedName.SubString($adComputer.DistinguishedName.IndexOf("OU=")))).DistinguishedName;
    #get pc onl in AD
    $computerad = get-adcomputer -Searchbase $ou -filter * -Properties operatingsystem | ? operatingsystem -match "windows" | Sort-Object name;
    $sourceFolder = "$LogFolder\PCTurnOnOvernight"
    $MainFolder = New-Item -Path "$sourceFolder" -ItemType Directory -Force;
    $testMainFolder = Test-Path -Path $MainFolder
    if ($testMainFolder -ne $true){
        $MainFolder
    }
    if (Test-Path -Path "$LogFolder\PCTurnOnOvernight\$getmonth"){
        Write-Host "Folder $month was created. Loading ..." -ForegroundColor Green;
    }else{
        New-Item -Path "$LogFolder\PCTurnOnOvernight\$getmonth" -ItemType Directory;
    }
    $filepath = "$sourceFolder\$getmonth\pconl_${getdate}.txt";
    $filepath1 = "$sourceFolder\pcoffl.txt";
    Out-File -FilePath $filepath;
    Out-File -FilePath $filepath1;
    Write-Host "Loading...Please wait" -ForegroundColor Yellow;
    $curent = 0;
    $countcomputer = $computerad.Count
    foreach ($cp in $computerad){
        $curent+=1;
        $cps = $cp.dnshostname;
        if (Test-Connection -ComputerName $cps -Quiet -Count 1){
            Add-Content -Path $filepath -Value $cp.Name -Force;
        }else {
            Add-Content -Path $filepath1 -Value $cp.Name -Force;
        }
        Write-Progress -Activity "Checking $cps" -Status "Loading ($curent/$countcomputer)" -PercentComplete ($current/$countcomputer*100)
    }
    $script:computeradonl = Get-Content -Path $filepath;
    $computeradonlnumber = $computeradonl.count;
    Write-Host "Total computer onine is $computeradonlnumber" -ForegroundColor Green;
}
function CheckPConline{
    param($hostcomputer,$getmonth,$credential,$computeradonl,$sourceFolder)
    $getdate = get-date -Format "dddd_MM_dd_yyyy";
    $folder = "Details_$getmonth"    
    $PathComputerOnl = "$sourceFolder\PCTurnOnOvernight\$getmonth\pconl_${getdate}.txt"
    $computeradonl = Get-Content -Path $PathComputerOnl
    $result = "$sourceFolder\PCTurnOnOvernight\$folder"
    $testDetailFolder = Test-Path -Path $result
    $current
    $countcomputer = $computeradonl.Count
    $getinfo = @();
    if ($testDetailFolder -ne $true){
        New-Item -Path "$sourceFolder\PCTurnOnOvernight\$folder" -ItemType "directory" -Force;
    }
    foreach ($pc in $computeradonl){
        $current += 1
        $newsession = New-PSSession -ComputerName $pc -Credential $credential -ErrorAction Ignore
        $scriptblock = {
            $computerinfo = $env:COMPUTERNAME
            $username = (Get-ComputerInfo).CSusername
            if ($username -eq $null){
                $getlocalmember = Get-LocalGroupMember Administrators | Where-Object {$_.ObjectClass -eq "User"} | Where-Object {$_.PrincipalSource -eq "ActiveDirectory"};
                $username = $getlocalmember.Name;
            }
            $details = @{
                "Date" = get-date -format MM/dd/yyyy
                "Computer name" = $computerinfo
                "Username" = $username
                "Status" = "Online"
            }
            New-Object psobject -Property $details 
            }
            $getinfo += Invoke-Command -Session $newsession -ScriptBlock $scriptblock -ErrorAction SilentlyContinue
            Write-Progress -Activity "Getting computer information of $pc" -Status "Loading ($current/$countcomputer)" -PercentComplete ($current/$countcomputer*100)
    }
    $getinfo | Export-Csv -Path "$result\detail_pcOnline_$getdate.csv" -NoTypeInformation   
}
#report each month 
function ReportEOM{
    param ($sourcePath,$getmonth)
    #delete PC offile file
    $path = "$sourcePath\PCTurnOnOvernight\$getmonth";
    if (Test-Path -Path "$path\report_${getmonth}.txt"){
        Remove-Item -Path "$path\report_${getmonth}.txt" -Force;
    }
    $getdate = (get-date -Format "m").split(" "); $getmonth = $getdate[0]
    $childitem = get-childitem -Path "$path" -Recurse;
    $childitemCount = $childitem.Count;
    #check file data and exit if there are no files remaining
    if ($childitemCount -eq 0){
        throw "Error! No data entry";
    }
    #sum and average
    $sum = 0;
    for ($i=0; $i -lt $childitemCount;$i++){
        $b = $childitem[$i]
        $Getcontent = Get-Content -Path "$path\$b" -Force;
        #exclude 60 pcs that was in use to build project
        $count = ($Getcontent.count-60);
        $sum = $sum + $count;
    }
    [int]$average = $sum/$childitemCount;
    Write-Host "------------------------------------------------------------------"
    Write-Host "Total computer left overnight in $getmonth is $sum in $childitemCount days" -ForegroundColor Green;
    Write-Host "------------------------------------------------------------------"
    Write-Host "Average number of machines left overnight daily in $getmonth is $average" -ForegroundColor Green;
    Write-Host "-------------------------------------------------------------------"
    Write-Warning "Check source folder ${sourcePath}source\ to more detailed information if something went wrong";
    Out-File -FilePath "$path\report_${getmonth}.txt"
    Add-Content -Path "$path\report_${getmonth}.txt" -Value "Total in month $getmonth $sum`nAverage one day is $average" -Force;
    Add-Content -Path "$path\report_${getmonth}.txt" -Value "Total days is $childitemCount" -Force;
    Invoke-Item "$path\report_${getmonth}.txt"
}
$hostcomputer = $env:COMPUTERNAME;
$getdate = get-date -Format "dddd_MM_dd_yyyy";
$month = (get-date -Format "m").split(" ");$getmonth = $month[0]
########################################################################
#this place to update variable. Update username and place to export report.
SetPasswordAutoSave -username "domain_admin"
$sourcePath = "D:\"
########################################################################
CreateLogFolder -LogFolder $sourcePath -getdate $getdate -getmonth $getmonth
CheckPConline -hostcomputer $hostcomputer -credential $credential -computeradonl $computeradonl -getmonth $getmonth -sourceFolder $sourcePath
Invoke-Item -Path "$sourcePath\PCTurnOnOvernight"
ReportEOM -sourcePath $sourcePath -getmonth $getmonth
