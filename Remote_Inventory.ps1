param (
    [string]$OU_Searchbase= "",  <# To perform a computer search on a domain #>
    [string]$DCIP="",  <# Specifies the IP of a DC to search for OUs #>
    [string]$IP = "127.0.0.1",  <# To perform a scan of a single IP #>
    [string]$w = "-wabefghipPqrsSu",  <# WMI options #>
    [switch]$a = $false,  <# Windows Installer Applications #>
    [switch]$b = $false,  <# BIOS Information #>
    [switch]$e = $false,  <# Event Log Files #>
    [switch]$f = $false,  <# File Shares #>
    [switch]$g = $false,  <# Local Groups (on non DC machines) #>
    [switch]$h = $false,  <# Additional Hardware #>
    [switch]$i = $false,  <# IP Routes #>
    [switch]$p = $false,  <# Printers #>
    [switch]$ps = $false,  <# Running Processes #>
    [switch]$q = $false,  <# Installed Patches #>
    [switch]$r = $false,  <# Registry Size #>
    [switch]$s = $false,  <# Services #>
    [switch]$st = $false,  <# Startup Commands #>
    [switch]$u = $false,  <# Local User accounts (on non DC machines) #>
    [string]$re = "-racdklp",  <# Registry Options #>
    [switch]$na = $false,  <# Non Windows Installer Applications #>
    [switch]$c = $false,  <# Windows Components #>
    [switch]$d = $false,  <# FQDN Domain Name #>
    [switch]$k = $false,  <# Product Keys #>
    [switch]$l = $false,  <# Last Logged on User #>
    [switch]$pr = $false,  <# Print Spooler Location #>
    [string]$user = "",  <# Username for Authentication #>
    [string]$pass = "",  <# Password for Authentication #>
    [switch]$ew = $false,  <# Export Word #>
    [switch]$ex = $false,  <# Export XML #>
    [switch]$et = $false,  <# Export TXT #>
    [string]$outPath = ".\remote_inventory_results.$outType", <# Specifies the results export location #>
    [string]$IPSubnet = "",  <# Subnet in xxx.xxx.xxx.xxx/xx form #>
    [string]$IPList = "",  <# A list of IPs to Scan seperated by commas #>
    [string]$IPDoc = "",  <# Path and file name of IPs #>
    [switch]$Help = $false  <# Displays help information #>
)

if ($Help -eq $true) {
    
    Write-Host "
    Usage: remote_inventory.ps1 [options]
    Examples: remote_inventory.ps1 -a -q -user user -pass password -IP 192.168.0.1 -ew
              remote_inventory.ps1 -a -q -user user -pass password -IPSubnet 192.168.0.0/24 -et
              remote_inventory.ps1 -a -q -user user -pass password -IPList 192.168.0.1, 192.168.0.2 -ex  
    ------------------------------------------------------------------------------------------------------
    Available Options:

    Specify Target Options:
    OU_Searchbase   - Specify an OU containing computers to scan. Must specify DCIP option
    DCIP            - Specify IP Address of the Domain Controller to Query
    IP              - Specify a single IP to scan
    IPSubnet        - Specify a subnet to scan. Must be in xxx.xxx.xxx.xxx/xx form. Currently
                      only available for /8, /16, and /24 subnets
    IPList          - Specify a list of IPs to scan. List must be seperated by commas
    IPDoc           - Specify a location and file name of a csv or txt document with a list of
                      IPs to scan
    
    Target Query Options:
    w               - Specify WMI options (Default: -wabefghipPqrsSu)
    a               - List Windows Installer Applications
    b               - Show BIOS Information
    e               - Show Location of Event Log Files
    f               - List File Shares
    g               - List Local Groups (on non DC machines)
    h               - List Additional Hardware (ie. Graphics Card)
    i               - List IP Routes
    p               - List Printers
    ps              - Show Running Processes
    q               - List Installed Patches
    r               - Show Registry Size
    s               - Show all Services
    st              - List Startup Commands
    u               - List Local User Accounts (on non DC machines)
    re              - Specify Registry Options (Default: -racdklp)
    na              - List Non Windows Installer Applications
    c               - List Windows Components
    d               - Show FQDN Domain Name
    k               - Show available Product Keys
    l               - Show Last Logged on User
    pr              - Show Print Spooler Location

    Authentication Options: 
    (NOTE: Authentication Options must be provided unless running against local machine)
    user            - Specify User Account to Authenticate 
    pass            - Specify Password to Authenticate

    Result Export Options:
    (NOTE: If no export option is selected, all outputs will be written to host only)
    ew              - Export to a Word Document
    ex              - Export to an XML Document
    et              - Export to a Text Document
    outPath         - Specify result output location and file name. i.e., \Temp\results.txt
                      Please remember to add the correct file type to end of file name.
    "

    exit
}

elseif ($IP -eq "127.0.0.1") {  <# if single IP is not specified #>

    if ($OU_Searchbase -ne "") {  <# Checks to see if OU_Searchbase was entered #>
        
        if ($DCIP -eq "") {
            Write-Host "DC IP not set. Exiting"
            exit
        }
        else {
            
        }
    <# need to create a query against the DC and then add the computers to a target array #>
    }

    elseif ($IPSubnet -ne "") {  <# Checks to see if IPSubnet was entered #>
    <# need to write a statement to convert a subnet to an array of targets #>
    }

    elseif ($IPList -ne "") {  <# Checks to see if IPList was entered #>
        $targets = $IPList.Split(",")
    }

    elseif ($IPDoc -ne "") {  <# Checks to see if a path and file was entered #>
        $file = $IPDoc.Split(".")
        $fileType = $file[1]

        if ($fileType -eq "csv") {

        }

        elseif ($fileType -eq "txt") {

        }

        else {
            Write-Host "File is not CSV or TXT format. File type is not supported"
            exit
        }
    
    }
}

else {
     $targets = $IP
}

ForEach ($target in $targets) {
    


}
