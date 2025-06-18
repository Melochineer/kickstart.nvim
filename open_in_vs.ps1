param (
    [string]$file,
    [int]$line,
    [int]$column = 1
)

# Normalize column to 1-based (minimum 1)
$column = [Math]::Max(1, $column)

# Start Visual Studio with file
Start-Process -FilePath "devenv" -ArgumentList "/edit $file" -NoNewWindow

# Retry DTE connection until success or timeout
$timeoutSeconds = 15
$retryIntervalMs = 500
$elapsed = 0
$dte = $null

while ($null -eq $dte -and $elapsed -lt $timeoutSeconds * 1000) {
    try {
        $dte = [System.Runtime.InteropServices.Marshal]::GetActiveObject("VisualStudio.DTE")
    }
    catch {
        Start-Sleep -Milliseconds $retryIntervalMs
        $elapsed += $retryIntervalMs
    }
}

if ($null -eq $dte) {
    Write-Error "Failed to connect to Visual Studio DTE after $timeoutSeconds seconds"
    exit 1
}

try {
    $doc = $dte.ActiveDocument
    if ($doc.FullName -ne $file) {
        $dte.ExecuteCommand("File.OpenFile", $file)
        Start-Sleep -Milliseconds 500
    }
    $activeDoc = $dte.ActiveDocument
    $textDoc = $activeDoc.Object("TextDocument")
    
    # Validate line number
    $endPoint = $textDoc.EndPoint
    $maxLines = $endPoint.Line
    $line = [Math]::Max(1, [Math]::Min($line, $maxLines))
    
    # Validate column number
    $startPoint = $textDoc.StartPoint.CreateEditPoint()
    $startPoint.MoveToLineAndOffset($line, 1)
    $lineEndPoint = $startPoint.CreateEditPoint()
    $lineEndPoint.EndOfLine()
    $lineLength = $lineEndPoint.AbsoluteCharOffset - $startPoint.AbsoluteCharOffset + 1
    $column = [Math]::Max(1, [Math]::Min($column, $lineLength))
    
    $selection = $activeDoc.Selection
    $selection.MoveToLineAndOffset($line, $column)
    $dte.MainWindow.Activate()
}
catch {
    Write-Error "Failed to set cursor position: $_"
    exit 1
}