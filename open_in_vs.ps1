param (
    [string]$file,
    [int]$line,
    [int]$column = 1
)

# Start Visual Studio with file
Start-Process -FilePath "devenv" -ArgumentList "/edit $file"

# Retry DTE connection until success or timeout
$timeoutSeconds = 3
$dte = $null
$activeDoc = $null
$selection = $null
$timer = [Diagnostics.Stopwatch]::StartNew()

while ($null -eq $dte -and $timer.elapsed.totalseconds -lt $timeoutSeconds)
{
    $dte = [System.Runtime.InteropServices.Marshal]::GetActiveObject("VisualStudio.DTE")
}

if ($null -eq $dte) {
    Write-Error "Failed to connect to Visual Studio DTE after $timeoutSeconds seconds"
    exit 1
}

$dte.MainWindow.Activate()
$timer = [Diagnostics.Stopwatch]::StartNew()

while ($null -eq $activeDoc -and $timer.elapsed.totalseconds -lt $timeoutSeconds)
{
    $activeDoc = $dte.ActiveDocument
}

if ($null -eq $activeDoc) {
    Write-Error "Failed to get to Active Document after $timeoutSeconds seconds"
    exit 1
}

$timer = [Diagnostics.Stopwatch]::StartNew()

while ($null -eq $selection -and $timer.elapsed.totalseconds -lt $timeoutSeconds)
{
	$selection = $activeDoc.Selection
}

if ($null -eq $selection) {
    Write-Error "Failed to get to Selection after $timeoutSeconds seconds"
    exit 1
}

$selection.MoveToLineAndOffset($line, $column)