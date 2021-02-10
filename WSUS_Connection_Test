$ServerType = Read-Host -Prompt "Is this an Upstream, Downstream, or Client?"

if ($ServerType -eq "Upstream" -or $ServerType -eq "upstream"){
    Write-Host "Testing connection from Upstream Server to Microsoft WSUS Server"

    $test1 = tnc windowsupdate.microsoft.com -Port 443 | Select-Object -Expand TcpTestSucceeded
    $test2 = tnc windowsupdate.microsoft.com -Port 80 | Select-Object -Expand TcpTestSucceeded

    if ($test1 -eq "True"){Write-Host "Success! Microsoft WSUS Server access over port 443 completed."}
    else {Write-Host "Failed. Microsoft WSUS Server access over port 443 failed."}

    if ($test2 -eq "True"){Write-Host "Success! Microsoft WSUS Server access over port 80 completed."}
    else {Write-Host "Failed. Microsoft WSUS Server access over port 80 failed."}

    if ($test1 -eq "True" -and $test2 -eq "True") {Write-Host "All tests completed successfully. Server able to connect to Microsoft WSUS Server over ports 443 and 80."}
    elseif ($test1 -eq "True" -and $test2 -eq "False") {Write-Host "Test to port 443 completed successfully, but test to port 80 failed. Please review firewall rules."}
    elseif ($test1 -eq "Fales" -and $test2 -eq "True") {Write-Host "Test to port 80 completed successfully, but test to port 443 failed. Please review firewall rules."}
    else{Write-Host "Unknown Failure. Please review above output to determine failed tests."}
}

elseif ($ServerType -eq "Downstream" -or $ServerType -eq "downstream") {
    $ServerIP = Read-Host -Prompt "Please input IP Address of Upstream Server"

    Write-Host "Testing connection from Downstream Server to Upstream Server"

    $test1 = tnc $ServerIP -Port 8530 | Select-Object -Expand TcpTestSucceeded
    $test2 = tnc $ServerIP -Port 8531 | Select-Object -Expand TcpTestSucceeded

    if ($test1 -eq "True"){Write-Host "Success! Upstream WSUS Server access over port 8530 completed."}
    else {Write-Host "Failed. Upstream WSUS Server access over port 8530 failed."}

    if ($test2 -eq "True"){Write-Host "Success! Upstream WSUS Server access over port 8531 completed."}
    else {Write-Host "Failed. Upstream WSUS Server access over port 8531 failed."}

    if ($test1 -eq "True" -and $test2 -eq "True") {Write-Host "All tests completed successfully. Server able to connect to the Upstream WSUS Server over ports 8530 and 8531."}
    elseif ($test1 -eq "True" -and $test2 -eq "False") {Write-Host "Test to port 8530 completed successfully, but test to port 8531 failed. Please review firewall rules."}
    elseif ($test1 -eq "Fales" -and $test2 -eq "True") {Write-Host "Test to port 8531 completed successfully, but test to port 8530 failed. Please review firewall rules."}
    else{Write-Host "Unknown Failure. Please review above output to determine failed tests."}
}

elseif ($ServerType -eq "Client" -or $ServerType -eq "client") {
    $ServerIP = Read-Host -Prompt "Please input IP Address of Downstream Server"

    Write-Host "Testing connection from Client to Downstream Server"

    $test1 = tnc $ServerIP -Port 8530 | Select-Object -Expand TcpTestSucceeded
    $test2 = tnc $ServerIP -Port 8531 | Select-Object -Expand TcpTestSucceeded

    if ($test1 -eq "True"){Write-Host "Success! Downstream WSUS Server access over port 8530 completed."}
    else {Write-Host "Failed. Downstream WSUS Server access over port 8530 failed."}

    if ($test2 -eq "True"){Write-Host "Success! Downstream WSUS Server access over port 8531 completed."}
    else {Write-Host "Failed. Downstream WSUS Server access over port 8531 failed."}

    if ($test1 -eq "True" -and $test2 -eq "True") {Write-Host "All tests completed successfully. Server able to connect to the Downstream WSUS Server over ports 8530 and 8531."}
    elseif ($test1 -eq "True" -and $test2 -eq "False") {Write-Host "Test to port 8530 completed successfully, but test to port 8531 failed. Please review firewall rules."}
    elseif ($test1 -eq "Fales" -and $test2 -eq "True") {Write-Host "Test to port 8531 completed successfully, but test to port 8530 failed. Please review firewall rules."}
    else{Write-Host "Unknown Failure. Please review above output to determine failed tests."}
}

else {Write-Host "Unknown Input. Please enter either 'Upstream', 'Downstream', or 'Client' only"}
