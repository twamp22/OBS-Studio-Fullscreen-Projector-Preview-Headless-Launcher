Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
#region Header ##########################################################
#																		#
#				OBS Studio Fullscreen Projector Preview					#
#							(headless launcher)							#
#																		#
#########################################################################
#																		#
#    This PowerShell script opens OBS Studio and attempts to connect    #
#    to it using an existing WebSocket. Once connected it launches a    #
#       fullscreen projector preview, showing the existing scene.       #
#																		#
#endregion###############################################################

$obsAddress = "localhost"
$obsPort = "4455"
$obsWebSocketAddress = "ws://$obsAddress`:$obsPort"
$obsExePath = "C:\Program Files\obs-studio\bin\64bit\obs64.exe"
Set-Location -Path "C:\Program Files\obs-studio\bin\64bit"
$d = $false  # Enable debugging messages?
$whichMonitor = 1  # Which monitor for OBS fullscreen preview (0, 1, 2, etc)
[System.Reflection.Assembly]::LoadFrom("WebSocket4Net.dll") | Out-Null

Write-Output "`nLaunching OBS fullscreen Projector!"

#region Functions
# Check for WebSocket
function CheckWebSocket {
    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Checking" }
    
    $tcpConnection = Test-NetConnection -ComputerName "localhost" -Port $obsPort
    
    return $tcpConnection.TcpTestSucceeded
}
# Wait for OBS WebSocket Server to be active
function WaitForWebSocket {
    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Waiting for WebSocket" }
    
    $maxA = 100  # 5 Seconds
    $a = 0
    
    while ($a -lt $maxA) {
        if (CheckWebSocket) {
            if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Active" }
            return  # Exit the loop immediately when WebSocket is active
        }
        Start-Sleep -Milliseconds 50
        $a++
    }

    Write-Output "[Error] [WEBSOCKET] Failed to connect in time, exiting."
    exit
}
# Check if OBS is running, if not, start it minimized
function RunOrCheckOBS {
    if ($d) { Write-Output "[DEBUG] [PROCESS] Checking OBS" }
    
    $obsRunning = Get-Process | Where-Object { $_.ProcessName -eq "obs64" }
    
    if (-not $obsRunning) {
        if ($d) { Write-Output "[DEBUG] [PROCESS] Starting OBS" }
        Start-Process -FilePath $obsExePath -ArgumentList "--minimize-to-tray --disable-safe-mode" -WindowStyle Minimized
        WaitForWebSocket  # Ensure WebSocket is active before continuing
    } else {
        if ($d) { Write-Output "[DEBUG] [PROCESS] OBS Running" }
    }

    # Now that OBS is confirmed running, attempt WebSocket connection
    ConnectToWebhook
}
# Connect to OBS WebSocket
function ConnectToWebhook {
    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Connecting" }
    
    $global:ws = New-Object WebSocket4Net.WebSocket($obsWebSocketAddress)
    $ws.Open()
    Start-Sleep -Milliseconds 200  # Give time to establish connection

    if ($ws.State -ne "Open") {
        Write-Output "[Error] [WEBSOCKET] Could not connect, exiting."
        exit
    }

    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Connected" }

    # Send Identify request
    $identifyRequest = @{
        op = 1
        d = @{
            rpcVersion = 1
        }
    } | ConvertTo-Json -Compress

    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] [IDENTITY] Sending request" }
    $ws.Send($identifyRequest)
    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] [IDENTITY] Sent:  $identifyRequest"}
    Start-Sleep -Milliseconds 100
    SendRequest
}
# Send Open Scene Projector request
function SendRequest {
    if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Sending request" }

    $obsCommand = @{
        op = 6
        d = @{
            requestType = "OpenVideoMixProjector"
            requestId = "1"
            requestData = @{
                videoMixType = "OBS_WEBSOCKET_VIDEO_MIX_TYPE_PROGRAM"
                monitorIndex = $whichMonitor
            }
        }
    } | ConvertTo-Json -Compress

    if ($ws -and $ws.State -eq "Open") {
        $ws.Send($obsCommand)
        if ($d) { Write-Output "[DEBUG] [WEBSOCKET] [REQUEST] Sent: $obsCommand" }
        Write-Output "OBS Fullscreen Projector Launched`n"
        $ws.Close()
        if ($d) { Write-Output "[DEBUG] [WEBSOCKET] Closed" }
    } else {
        Write-Output "[Error] [WEBSOCKET] [REQUEST] Connection lost before sending request."
    }
}
#endregion

#region Main
RunOrCheckOBS  # Start OBS if needed, then connect

#endregion
