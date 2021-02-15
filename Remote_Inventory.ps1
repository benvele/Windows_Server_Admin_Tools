param (
    [string]$IP = "127.0.0.1",  <# To perform a scan of a single IP #>
    [string]$IPSubnet = "", <# Specify Subnet to scan #>
    [string]$OU_Searchbase = "", <# Specify which OU to search for computers #>
    [string]$DCIP = "", <# Specify IP of DC for OU Search #>
    [switch]$a = $false,  <# Applications #>
    [switch]$q = $false,  <# Installed Patches #>
    [switch]$b = $false, <# BIOS Information #>
    [switch]$e = $false, <# Last 50 Events from System Event Log #>
    [switch]$f = $false, <# File Shares #>
    [switch]$g = $false, <# Local Groups #>
    [switch]$h = $false, <# Additional Hardware #>
    [switch]$i = $false, <# IP Routes #>
    [switch]$p = $false, <# List Printers #>
    [switch]$ps = $false, <# List Running Processes #>
    [switch]$r = $false, <# Registry Size #>
    [switch]$s = $false, <# Services #>
    [switch]$st = $false, <# Startup Commands #>
    [switch]$u = $false, <# Local User Accounts #>
    [switch]$c = $false, <# Windows Components #>
    [switch]$d = $false, <# FQDN #>
    [switch]$k = $false, <# Product Keys #>
    [switch]$l = $false, <# Last Logged in User #>
    [switch]$pr = $false, <# Print Spooler Location #>
    [string]$user = "",  <# Username for Authentication #>
    [string]$pass = "",  <# Password for Authentication #>
    [switch]$ew = $false,  <# Export Word #>
    [switch]$ec = $false,  <# Export CSV #>
    [switch]$et = $false,  <# Export TXT #>
    [switch]$ex = $false,  <# Export XML #>
    [string]$outPath = "", <# Specifies the results export location #>
    [string]$IPList = "",  <# A list of IPs to Scan seperated by commas #>
    [string]$IPDoc = ".\",  <# Path and file name of IPs #>
    [switch]$Help = $false  <# Displays help information #>
)

$localHost = $true
$outType = "csv"

if ($Help -eq $true) {
    
    Write-Host "
    Usage: remote_inventory.ps1 [options]
    Examples: remote_inventory.ps1 -a -q -user user -pass password -IP 192.168.0.1 -ew
              remote_inventory.ps1 -a -q -user user -pass password -IPSubnet 192.168.0.0/24 -et
              remote_inventory.ps1 -a -q -user user -pass password -IPList 192.168.0.1, 192.168.0.2 -ex  
    ------------------------------------------------------------------------------------------------------
    Available Options:

    Specify Target Options: (Only one can be specified)
    IP              - Specify a single IP to scan
    IPList          - Specify a list of IPs to scan. List must be seperated by commas
    IPDoc           - Specify a location and file name of a csv or txt document with a list of
                      IPs to scan. Must be a single line seperated by commas
    OU_Searchbase   - Specify OU to search for computer accounts. Requires DCIP option to be set
    DCIP            - IP Address of DC to be used with OU_Searchbase option
    
    Target Query Options:
    a               - List Applications
    q               - List Installed Patches
    b               - Display BIOS Information
    e               - Show Last 50 Events from System Event Log
    f               - List File Shares
    g               - List Local User Groups
    h               - List Additional Hardware
    i               - Show IP Routes
    p               - List Installed Printers
    ps              - Show Running Processes
    r               - Display Registry Size
    s               - Display all Services
    st              - List Startup Commands
    u               - List Local User Accounts
    c               - List Windows Components
    d               - Show FQDN
    k               - List Available Product Keys
    l               - Show Last Logged on User
    pr              - Display Print Spooler Location

    Authentication Options: 
    (NOTE: Authentication Options must be provided unless running against local machine)
    user            - Specify User Account to Authenticate. Must be in DOMAIN\Username format 
    pass            - Specify Password to Authenticate

    Result Export Options:
    (NOTE: If no export option is selected, all outputs will be written to host only)
    ew              - Export to a Word Document
    ec              - Export to an CSV Document
    et              - Export to a Text Document
    ex              - Export to an XML Document
    outPath         - Specify result output location and file name. i.e., \Temp\results.txt
                      Please remember to add the correct file type to end of file name.
    "

    exit
}

elseif ($IP -eq "127.0.0.1" ) {  <# if single IP is not specified #>


    if ($IPList -ne "") {  <# Checks to see if IPList was entered #>
        $targets = $IPList.Split(",").Trim()
        $localHost = $false
    }

    elseif ($IPDoc -ne ".\") {  <# Checks to see if a path and file was entered #>
        $file = $IPDoc.Split(".")
        $fileType = $file[$file.Count-1]

        if ($fileType -eq "csv" -or $fileType -eq "txt") {
            $hosts = Get-Content $IPDoc
            $targets = $hosts.Split(",").Trim()
        }

        elseif($IPDoc -eq ".\"){}

        else {
            Write-Host "File is not CSV or TXT format. File type is not supported"
            exit
        }
        $localHost = $false
    }

    elseif ($IPSubnet -ne "") {

        $subnetParts = $IPSubnet.Split('/')

        $octets = $subnetParts[0].Split('.')
        $mask = $subnetParts[1]

        [int]$w = $octets[0]
        [int]$x = $octets[1]
        [int]$y = $octets[2]
        [int]$z = $octets[3]

        $countMask = (32 - $mask)

        $count = [Math]::Pow(2, $countMask)

        for ($count -lt 0; $count--) {
    
            $targets = $targets += "$w.$x.$y.$z"
            $z++

            if ($z -gt 255) {
                $y++
                $z=0
            }      
    
            if ($y -gt 255) {
                $x++
                $y=0
            }
            if ($x -gt 255) {
                $w++
                $x=0
            }
        
        }
    }

    else {
        $targets = $IP
        $localHost = $true
    }
}
else {
    $localHost = $false
}

if ($ec -eq $true) {$outType = "csv"}
if ($et -eq $true) {$outType = "txt"}
if ($ew -eq $true) {
    $Word = New-Object -ComObject Word.Application
    $Document = $Word.Documents.Add()
    $Selection = $Word.Selection
    $outType = "docx"
}
$workingPath = pwd | Select-Object -expand Path

if($outPath -eq "") {
    $outPath = $workingPath + "\remote_inventory_results.$outType"
}
$topRow = '"Target:"'

if ($a -eq $true) {$topRow = $topRow +',"Loaded Apps:"'}
if ($q -eq $true) {$topRow = $topRow + ',,,"Installed Patches:"'}
if ($b -eq $true) {$topRow = $topRow + ',,"BIOS Information:"'}
if ($e -eq $true) {$topRow = $topRow + ',"Event Log Entries:"'}

if ($ec -eq $true -or $et -eq $true){Add-Content -Path $outPath -Value $topRow}
elseif($ew -eq $true) {
    $Selection.Style = 'Title'
    $Selection.TypeText('Remote Inventory Results:')
    $Selection.TypeParagraph()
    $Selection.Font.Bold = 1
    $Selection.Font.Italic = 1
    $Selection.TypeText("Targets will be in Heading 2 Format. Scroll or use Navigation Pane to navigate through targets.")
    $Selection.Font.Bold = 0
    $Selection.Font.Italic = 0
    $Selection.TypeParagraph()
}
else {Write-Host $topRow}

if ($localHost -eq $true) {
    ForEach ($target in $targets) {
        $row = "$target"

        if ($a -eq $true) {
            $loadedApps = ""
            $wordAppContent = @()

            $apps = Get-Package | Select-Object -expand Name
            ForEach ($app in $apps) {
                $loadedApps = $loadedApps + "$app; "
                $wordAppContent += $app
            
            }
            $row = $row + ",$loadedApps"
        }

        if ($q -eq $true) {
            $installedPatches = ""
            $wordPatchContent = @()

            $patches = Get-WmiObject -Class win32_quickfixengineering | Select-Object -expand HotFixID
            ForEach ($patch in $patches) {
                $installedPatches = $installedPatches + "$patch; "
                $wordPatchContent += $patch
            }
            $row = $row + ",$installedPatches"
        }

        if ($b -eq $true) {
            $biosInfo = ""
            $wordBiosInfo = @()

            $bios = Get-WmiObject -Class Win32_BIOS

            $biosInfo = "Computer Name: " + $bios.PSComputerName + " ; Manufacturer: " + $bios.Manufacturer + " ; Serial Number: " + $bios.SerialNumber + " ; BIOS Version: " + $bios.SMBIOSBIOSVersion
            $wordBiosInfo += "Computer Name: " + $bios.PSComputerName + " ; Manufacturer: " + $bios.Manufacturer + " ; Serial Number: " + $bios.SerialNumber + " ; BIOS Version: " + $bios.SMBIOSBIOSVersion

            $row = $row + ",$biosInfo"
        }

        if ($e -eq $true) {
            $events = ""
            $wordEvents = @()

            $logs = Get-EventLog -LogName System -Newest 50

            foreach ($log in $logs) {
                [string]$entry = "Time: " + [string]$log.TimeGenerated + ", Entry Type: " + $log.EntryType + ", Instance ID: " + $log.InstanceId + ", Message: " + $log.Message + ' ||| '
                $events = $events + $entry
                $wordEvents += $entry
            }

            $row = $row + ",$events"

        }

    }
}

if ($localHost -eq $false) {
    if ($user -eq "" -or $pass -eq ""){
        Write-Host "Username or Password not supplied. Please run again and supply missing credential"
        if ($ew -eq $true) {
            $word.Quit()
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            Remove-Variable word 
        }
        exit
    }
    else {
        $Password = ConverrTo-SecureString $pass -AsPlainTest -Force
        $creds = New-Object System.Management.Automation.PSCredential $user, $Password
    }

    ForEach ($target in $targets) {
        $row = "$target"

        if ($a -eq $true) {
            $loadedApps = ""
            $wordAppContent = @()

            $apps = Invoke-Command -ComputerName $target -Credential $creds -ScriptBlock {Get-Package | Select-Object -expand Name}
            ForEach ($app in $apps) {
                $loadedApps = $loadedApps + "$app; "
                $wordAppContent += $app
            
            }
            $row = $row + ",$loadedApps"
        }

        if ($q -eq $true) {
            $installedPatches = ""
            $wordPatchContent = @()

            $patches = Get-WmiObject -ComputerName $target -Credential $creds -Class win32_quickfixengineering | Select-Object -expand HotFixID
            ForEach ($patch in $patches) {
                $installedPatches = $installedPatches + "$patch; "
                $wordPatchContent += $patch
            }
            $row = $row + ",$installedPatches"
        }

        if ($b -eq $true) {
            $biosInfo = ""
            $wordBiosInfo = @()

            $bios = Get-WmiObject -ComputerName $target -Credential $creds -Class Win32_BIOS
            
            $biosInfo = "Computer Name: " + $bios.PSComputerName + "; Manufacturer: " + $bios.Manufacturer + "; Serial Number: " + $bios.SerialNumber + "; BIOS Version: " + $bios.SMBIOSBIOSVersion
            $wordBiosInfo += "Computer Name: " + $bios.PSComputerName + "; Manufacturer: " + $bios.Manufacturer + "; Serial Number: " + $bios.SerialNumber + "; BIOS Version: " + $bios.SMBIOSBIOSVersion

            $row = $row + ",$biosInfo"
        }

        if ($e -eq $true) {
            $events = ""
            $wordEvents = @()

            $logs = Invoke-Command -ComputerName $target -Credential $creds -ScriptBlock {Get-EventLog -ComputerName $target -LogName System -Newest 50}

            foreach ($log in $logs) {
                [string]$entry = "Time: " + [string]$log.TimeGenerated + ", Entry Type: " + $log.EntryType + ", Instance ID: " + $log.InstanceId + ", Message: " + $log.Message + ' ||| '
                $events = $events + $entry
                $wordEvents += $entry
            }

            $row = $row + ",$events"

        }
    }
}

if ($ec -eq $true -or $et -eq $true){Add-Content -Path $outPath -Value $row}
elseif($ew -eq $true){
    $Selection.TypeParagraph()
    $Selection.Style = 'Heading 2'
    $Selection.TypeText($target)
    $Selection.TypeParagraph()
    if($a -eq $true){
        $Selection.Style = 'Heading 3'
        $Selection.TypeText("Installed Apps:")
        $Selection.TypeParagraph()
        forEach ($inApp in $wordAppContent){
            $Selection.TypeText($inApp)
            $Selection.TypeParagraph()
        }
    }
    if($q -eq $true) {
        $Selection.TypeParagraph()
        $Selection.Style = 'Heading 3'
        $Selection.TypeText("Installed Patches:")
        $Selection.TypeParagraph()
        forEach ($inPatch in $wordPatchContent) {
            $Selection.TypeText($inPatch)
            $Selection.TypeParagraph()
        }
    }
    if($b -eq $true) {
        $Selection.TypeParagraph()
        $Selection.Style = 'Heading 3'
        $Selection.TypeText("BIOS Information:")
        $Selection.TypeParagraph()
        $Selection.TypeText($wordBiosInfo)
        $Selection.TypeParagraph()
    }
    if($e -eq $true) {
        $Selection.TypeParagraph()
        $Selection.Style = 'Heading 3'
        $Selection.TypeText("Last 50 System Events:")
        $Selection.TypeParagraph()
        forEach ($event in $wordEvents){
            $Selection.TypeText($event)
            $Selection.TypeParagraph()
        }
    }
    $Document.SaveAs($outPath)
    $word.Quit()

    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable word 
}
else{Write-Host $row}

