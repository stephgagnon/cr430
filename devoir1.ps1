# Rendre la taille d'un fichier lisible pour l'humain et l'arrondir

function Get-FormattedFileSize
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$true,Position=0)]
        [ulong] $fileSize
    )

    switch($fileSize) {
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

# Obtenir les ordinateurs actifs dans Active Directory

$vmNames = Get-ADComputer -Filter * -Properties Name,LastLogonDate | Where-Object {$_.Enabled} | Select-Object -ExpandProperty Name
$offlineVMs = @()

# Création du rapport

$reportSections = @()

ForEach ($vmName in $vmNames) {
    
    $vmSession = New-PSSession $vmName -ErrorAction SilentlyContinue

    if ($null -eq $vmSession) {
        Write-Host "$vmName is offline."
        $offlineVMs += $vmName
        Continue;
    }

    $reportSections += "<h1>Computer name: $vmName</h1>"
    
    # Obtenir les informations du système d'exploitation, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $osInfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_OperatingSystem
    } -Session $vmSession

    $uptimeFilter = @{Label='Uptime';Expression={((Get-Date) - $_.LastBootUpTime).ToString("hh\:mm\:ss")}}
    $osInfoHTML = $osInfo | ConvertTo-Html -Property Version,Caption,BuildNumber,$upTimeFilter -Fragment -PreContent "<h2>Operating System Information</h2>"
    $reportSections += $osInfoHTML

    # Obtenir les informations du processeur, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $processInfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_Processor
    } -Session $vmSession

    $processInfoHTML = $processInfo | ConvertTo-Html -Property DeviceID,Name,Caption -Fragment -PreContent "<h2>Processor Information</h2>"
    $reportSections += $processInfoHTML

    # Obtenir les informations du disque, appeler la fonction Get-FormattedFileSize, convertir le résultat en code HTML dans une table et le stocker dans une variable

    $diskInfo = Invoke-Command -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    } -Session $vmSession

    $sizeFilter = @{Label='Size';Expression={Get-FormattedFileSize($_.Size)}}
    $freeSpaceFilter = @{Label='FreeSpace';Expression={Get-FormattedFileSize($_.FreeSpace)}}
    $diskInfo = $diskInfo | Select-Object -Property DeviceID,DriveType,$sizeFilter,$freeSpaceFilter
    $diskInfoHTML = $diskInfo | ConvertTo-Html -Property DeviceID,DriveType,Size,FreeSpace -Fragment -PreContent "<h2>Disk Information</h2>"
    $reportSections += $diskInfoHTML
}

# Toutes les informations sont rassemblées dans un seul rapport HTML, incluant la date de création du rapport

$reportSections += "<h2>Offline Computers: $($offlineVMs | Join-String -Separator ", ")</h2>"
$reportBody = $reportSections | Join-String -Separator " "
$report = ConvertTo-HTML -Body $reportBody -Title "Computer Information Report" -PostContent "<p><b>Date:<b> $((Get-Date).ToString("dd-MM-yyyy"))<p>"

# Générer le rapport en un fichier HTML

$reportName = "ComputerReport_" + (Get-Date -Format "ddMMyyyyy") + ".html"
$report | Out-File .\$reportName