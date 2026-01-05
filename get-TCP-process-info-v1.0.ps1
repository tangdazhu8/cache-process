# 提权
$isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $isElevated) {
	Write-Host "Restarting with elevated permissions..."
	Start-Sleep -Seconds 2
	$scriptPath = $MyInvocation.MyCommand.Path
	$newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
	$newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
	$newProcess.Verb = "runas"
	[System.Diagnostics.Process]::Start($newProcess) | Out-Null
	exit
}

# verify port
function verify-port {
	param([string]$InputPort)
	if ($InputPort -match '^\d+$') {
		$number = [int]$InputPort
		if ($number -ge 1 -and $number -le 65535) {
			return $true
		}
	}
	return $false
}

# create ProcessInfo file on desktop
$timestamp = get-date -format "yyMMdd"
$outputFile = "$env:userprofile\desktop\TCP_ProcessInfo_$timestamp.txt"

# add generate time
"Generated: $(get-date -format 'yyyy-MM-dd HH:mm:ss')" | out-file -filepath $outputFile -Append -Encoding UTF8
"=" * 50 | out-file -filepath $outputFile -Append -Encoding UTF8
"" | out-file -filepath $outputFile -Append -Encoding UTF8

# main
do {
	
	# input TCP port
	$tcp_port = Read-Host "please input TCP port"
	do {
		if (verify-port -InputPort $tcp_port) {
			break
		} else {
			$tcp_port = Read-Host "Invalid input. Please input TCP port again"
		}
	} while ($true)
	
	# get pid
	$get_pid = (netstat -ano 2>$null | Select-String "TCP.*:$tcp_port\s+" | Select-Object -First 1) -replace '.*\s+(\d+)$', '$1'
	if (-not $get_pid) {
		"The TCP port $tcp_port was not found."
		continue
	}
	
	# check the process exists
	$process = Get-Process -Id $get_pid -ErrorAction SilentlyContinue
	if (-not $process) {
		"The process PID $get_pid was not found."
		continue
	}
	
	# get process name
	$processname = (Get-Process -Id $get_pid).processname
	
	# get system.exe version from kernel file "ntoskrnl.exe"
	if ($processname -eq "system") {
		$filepath = "$env:windir\system32\ntoskrnl.exe"
		if (test-path $filepath) {
			$info = (get-item $filepath).versionInfo
		} else {
			$info = $null
		}
		
	# get system core process version from file path
	} elseif ($processname -in @("lsass", "wininit", "services")) {
		$filepath = "$env:windir\system32\$processname.exe"
		if (test-path $filepath) {
			$info = (get-item $filepath).Versioninfo
		} else {
			$info = $null
		}
		
	# get non-system core process from mainmodule
	} else {
		if ($process.mainmodule) {
				$info = $process.MainModule.FileVersionInfo
		} else {
				$info = $null
		} 
	}
	
	# get service
	$services = (Get-CimInstance Win32_Service | Where-Object {$_.ProcessId -eq $get_pid})
	if ($services) {
		$service = $services.name
	} else {
		$service = "N/A"
	}
	
	# get version
	if ($info) {
		$accurateVersion = "{0}.{1}.{2}.{3}" -f $info.FileMajorPart, $info.FileMinorPart, $info.FileBuildPart, $info.FilePrivatePart
	} else {
		$accurateVersion = "N/A"
	}
	
	# record information
	$result = "
	TCP port:   $tcp_port
	PID:        $get_pid
	process:    $processname
	service:    $service
	version:    $accurateVersion
	"
	
	# output
	$result | Out-File -filepath $outputFile -append -Encoding UTF8
	
} while ($true)
