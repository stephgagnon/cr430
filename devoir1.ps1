# Rendre la taille d'un fichier lisible pour l'humain et l'arrondir

function Get-FormattedFileSize
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$true,Position=0)]
        [ulong] $FileSize
    )

    switch($FileSize) {
        {$_ -gt 1TB }
            { "{0:n2} TB" -f ($_ / 1TB); break }
        {$_ -gt 1GB }
            { "{0:n2} GB" -f ($_ / 1GB); break }
        {$_ -gt 1MB }
            { "{0:n2} MB" -f ($_ / 1MB); break }
        {$_ -gt 1KB }
            { "{0:n2} KB" -f ($_ / 1KB); break }
        default
            { "{0:n2} B" -f $_ }
    }
}

# Faciliter la création d'un objet PSCredential

function Get-PSCredential
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$true,Position=0)]
        [string] $UserName,
        [string] $Password
    )

    if ([string]::IsNullOrEmpty($Password)) {
    	$Credential = Get-Credential -UserName $UserName
    	if ($PSEdition -eq 'Desktop') {
	        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
	        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
	    } else {
	        $Password = ConvertFrom-SecureString -SecureString $SecureString -AsPlainText
	    }
    } else {
	    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
		$Credential = New-Object System.Management.Automation.PSCredential @($UserName, $SecurePassword)
    }

    $Credential
}

# Définitions pour le PowerShell à distance

$VMNames = @("IT-HELP-DC","IT-HELP-GW","IT-HELP-DVLS","IT-HELP-WAC")
$UserName = "IT-HELP\Administrator"
#$Password = "DevoLabs123!"
#$Credential = Get-PSCredential -UserName $UserName -Password $Password

# Création du rapport

$ReportSections = @()

ForEach ($VMName in $VMNames) {
    if($Credential){
        $VMSession = New-PSSession -VMName $VMName -Credential $Credential
    } else {
        $VMSession = New-PSSession $VMName    
    }
    
    # Obtenir le nom de la machine
    $NameInfo = Invoke-Command -ScriptBlock {
        [PSCustomObject]@{
            ComputerName = $env:computername
        }
    } -Session $VMSession
    $NameInfoHTML = "<h1>Computer name: $($NameInfo.ComputerName)</h1>"
    $ReportSections += $NameInfoHTML

    # Obtenir les informations du système d'exploitation, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $OSinfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_OperatingSystem
    } -Session $VMSession

    $OSinfoHTML = $OSinfo | ConvertTo-Html -Property Version,Caption,BuildNumber -Fragment -PreContent "<h2>Operating System Information</h2>"
    $ReportSections += $OSinfoHTML

    # Obtenir les informations du processeur, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $ProcessInfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_Processor
    } -Session $VMSession

    $ProcessInfoHTML = $ProcessInfo | ConvertTo-Html -Property DeviceID,Name,Caption -Fragment -PreContent "<h2>Processor Information</h2>"
    $ReportSections += $ProcessInfoHTML

    # Obtenir les informations du disque, appeler la fonction Get-FormattedFileSize, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $DiskInfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    } -Session $VMSession

    $SizeFilter = @{Label='Size';Expression={Get-FormattedFileSize($_.Size)}}
    $FreeSpaceFilter = @{Label='FreeSpace';Expression={Get-FormattedFileSize($_.FreeSpace)}}
    $DiskInfo = $DiskInfo | Select-Object -Property DeviceID,DriveType,$SizeFilter,$FreeSpaceFilter
    $DiskInfoHTML = $DiskInfo | ConvertTo-Html -Property DeviceID,DriveType,Size,FreeSpace -Fragment -PreContent "<h2>Disk Information</h2>"
    $ReportSections += $DiskInfoHTML
}

# Toutes les informations sont rassemblées dans un seul rapport HTML, incluant la date de création du rapport

$ReportBody = $ReportSections | Join-String -Separator " "
$Report = ConvertTo-HTML -Body $ReportBody -Title "Computer Information Report" -PostContent "<p><b>Date:<b> $((Get-Date).ToString("dd-MM-yyyy"))<p>"

# Générer le rapport en un fichier HTML
$Report | Out-File .\ComputerReport.html