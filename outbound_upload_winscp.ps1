<#
.SYNOPSIS
    Enterprise-grade SFTP file upload solution designed by Sumit
.DESCRIPTION
    Uploads XML files to MEdius SFTP server using WinSCP.
.NOTES
    Author: Sumit Choudhary
    Version: 3.3
    Last Updated: 2025-03-14
#>

# Define variables
$WinSCPPath       = "C:\Program Files (x86)\WinSCP\WinSCP.com"
$SFTPHost         = "sftp.rei.ricoh.com"
$SFTPUser         = "medius_rei"
$PrivateKeyPath   = "E:\SFTP_KEY\medius_rei.ppk"
$HostKey          = "ssh-rsa 4096 1hUCTS8g2QB5iAQLlCKfBf/xp8ntEx5RVljeddJAhUE"
$LocalDirectory   = "D:\REI_Medius\PRG\Supplier\MEDIUS_OUTBOUND\outbound"
$RemoteDirectory  = "/medius_fx/outbound/"
$ArchiveDirectory = "D:\REI_Medius\PRG\Supplier\ARC_OUT"
$LogDirectory     = "D:\REI_Medius\PRG\Supplier\MEDIUS_OUTBOUND\logs"
$SupplierArcInDir = "D:\REI_Medius\PRG\Supplier\ARC_IN"
# Updated to use the HTML email sending script
$VBScriptPath     = "D:\REI_Medius\UTL\SendEmail2Admin_HTML.vbs"
$EmailSubject     = "Files Uploaded to SFTP"

# Create timestamp for log files
$Timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile     = "$LogDirectory\upload_log_$Timestamp.txt"
$SFTPLogFile = "$LogDirectory\sftp_log_$Timestamp.txt"

# HTML email header and footer for a consistent and modern look
$HTMLHeader = @"
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, Helvetica, sans-serif; background-color: #f9f9f9; color: #333; }
    h2 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 5px; }
    p { line-height: 1.4; }
    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
    th { background-color: #0078d4; color: #fff; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .success { color: #008000; font-weight: bold; }
    .failed { color: #ff0000; font-weight: bold; }
    .unknown { color: #ff8c00; font-weight: bold; }
    .supplier-info { background-color: #e6f7ff; padding: 10px; border-left: 4px solid #0078d4; margin: 10px 0; }
    .summary { font-size: 14px; }
  </style>
</head>
<body>
"@

$HTMLFooter = @"
</body>
</html>
"@

# Function to write to log file and console
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Output $logMessage
    Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
}

# Function to extract supplier code from XML file
function Get-SupplierCodeFromXML {
    param (
        [string]$FilePath
    )

    try {
        [xml]$xml = Get-Content -Path $FilePath -Encoding UTF8
        if ($xml.VendorCode -and $xml.VendorCode.SupplierCode) {
            return $xml.VendorCode.SupplierCode
        }
        return $null
    }
    catch {
        Write-Log "ERROR: Failed to extract supplier code from $FilePath. Error: $($_.Exception.Message)"
        return $null
    }
}

# Function to find supplier name from ARC_IN directory based on supplier code
function Get-SupplierNameFromCode {
    param (
        [string]$SupplierCode
    )

    try {
        $supplierFiles = Get-ChildItem -Path $SupplierArcInDir -Filter "*_${SupplierCode}_*.xml" -File
        
        foreach ($file in $supplierFiles) {
            try {
                [xml]$xml = Get-Content -Path $file.FullName -Encoding UTF8
                if ($xml.Supplier -and $xml.Supplier.SupplierName -and $xml.Supplier.SupplierCode -eq $SupplierCode) {
                    return @{
                        Name = $xml.Supplier.SupplierName
                        File = $file.Name
                    }
                }
            }
            catch {
                Write-Log "WARNING: Could not parse file $($file.FullName). Error: $($_.Exception.Message)"
                continue
            }
        }
        return $null
    }
    catch {
        Write-Log "ERROR: Failed to search for supplier name. Error: $($_.Exception.Message)"
        return $null
    }
}

# Ensure directories exist
foreach ($dir in @($LogDirectory, $ArchiveDirectory)) {
    if (!(Test-Path -Path $dir)) {
        try {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Log "Created directory: $dir"
        }
        catch {
            Write-Output "ERROR: Failed to create directory $dir. Error: $($_.Exception.Message)"
            exit 1
        }
    }
}

# Initialize log file
$LogContent = @"
SFTP Upload Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
========================================
Local Directory: $LocalDirectory
Remote Directory: $RemoteDirectory
SFTP Host: $SFTPHost
SFTP User: $SFTPUser
========================================

"@

try {
    Set-Content -Path $LogFile -Value $LogContent -Encoding UTF8
}
catch {
    Write-Output "ERROR: Failed to create log file. Error: $($_.Exception.Message)"
    exit 1
}

# Get all files in the local directory
try {
    $Files = Get-ChildItem -Path $LocalDirectory -File
    Write-Log "Found $($Files.Count) files in $LocalDirectory"
}
catch {
    Write-Log "ERROR: Failed to access directory $LocalDirectory. Error: $($_.Exception.Message)"
    exit 1
}

# Check if there are files to upload
if ($Files.Count -eq 0) {
    $message = "No files found in $LocalDirectory to upload."
    Write-Log $message
    
    # Build HTML email notifying that no files were found
    $HTMLMessage = $HTMLHeader + @"
<h2>No Files to Upload</h2>
<p>No files were found in <strong>$LocalDirectory</strong> to upload to the SFTP server.</p>
<p>Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
"@ + $HTMLFooter
    
    $WScriptPath = "$Env:SystemRoot\System32\cscript.exe"
    try {
        & $WScriptPath $VBScriptPath "$EmailSubject - No Files" "" "$HTMLMessage" 2>&1 | Out-Null
        Write-Log "Email notification sent about no files to upload."
    }
    catch {
        Write-Log "ERROR: Failed to send email notification. Error: $($_.Exception.Message)"
    }
    exit 0
}

# Collect supplier information for each file
$FilesWithSupplierInfo = @()
foreach ($File in $Files) {
    $supplierCode = Get-SupplierCodeFromXML -FilePath $File.FullName
    $supplierInfo = $null
    
    if ($supplierCode) {
        $supplierInfo = Get-SupplierNameFromCode -SupplierCode $supplierCode
        Write-Log "File $($File.Name) - Supplier Code: $supplierCode, Supplier Name: $(if ($supplierInfo) { $supplierInfo.Name } else { 'Not Found' })"
    }
    else {
        Write-Log "File $($File.Name) - Could not extract supplier code"
    }
    
    $FilesWithSupplierInfo += [PSCustomObject]@{
        File = $File
        SupplierCode = $supplierCode
        SupplierName = if ($supplierInfo) { $supplierInfo.Name } else { "Unknown" }
        SourceFile = if ($supplierInfo) { $supplierInfo.File } else { "N/A" }
    }
}

# Log the files to be uploaded
Write-Log "Files to be uploaded:"
foreach ($FileInfo in $FilesWithSupplierInfo) {
    $fileDetails = "- $($FileInfo.File.Name) (Size: $($FileInfo.File.Length) bytes, Supplier: $($FileInfo.SupplierName), Code: $($FileInfo.SupplierCode))"
    Write-Log $fileDetails
}
Write-Log ""

# Create a temporary WinSCP script file
$ScriptPath = Join-Path $env:TEMP "WinSCPUploadScript_$Timestamp.txt"

# Create WinSCP script with explicit connection parameters and disabled timestamp and permission preservation.
$WinSCPCommands = @(
    "# Disable interactive prompts",
    "option batch on",
    "option confirm off",
    "",
    "# Disable preserving timestamp and permissions",
    "option transfer preserve timestamp off",
    "option transfer preserve permissions off",
    "",
    "# Connect using explicit parameters instead of saved session",
    "open sftp://$SFTPUser@$SFTPHost/ -hostkey=`"$HostKey`" -privatekey=`"$PrivateKeyPath`" -rawsettings AuthKIPassword=0 AuthGSSAPI=0",
    "",
    "# Change to target directory",
    "cd $RemoteDirectory",
    ""
)

# Add each file to be uploaded
foreach ($FileInfo in $FilesWithSupplierInfo) {
    $LocalFilePath = $FileInfo.File.FullName.Replace('\', '\\')
    $WinSCPCommands += "put `"$LocalFilePath`""
}

$WinSCPCommands += "exit"

# Write script to file
try {
    $WinSCPCommands | Out-File $ScriptPath -Encoding ASCII
    Write-Log "WinSCP script created at: $ScriptPath"
}
catch {
    Write-Log "ERROR: Failed to create WinSCP script. Error: $($_.Exception.Message)"
    exit 1
}

# Log the start of upload process
$startMessage = "Starting upload process at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')..."
Write-Log $startMessage

# Execute WinSCP with the script file
try {
    $process = Start-Process -FilePath $WinSCPPath -ArgumentList "/script=`"$ScriptPath`"", "/log=`"$SFTPLogFile`"" -NoNewWindow -Wait -PassThru
    $exitCode = $process.ExitCode
    
    $exitCodeMessage = "WinSCP process completed with exit code: $exitCode"
    Write-Log $exitCodeMessage
}
catch {
    $errorMessage = "Error executing WinSCP: $($_.Exception.Message)"
    Write-Log $errorMessage
    $exitCode = 99
}

# Process upload results
$uploadResults = @()
$uploadSuccess = $false

# Check if SFTP log exists and parse it
if (Test-Path $SFTPLogFile) {
    try {
        $logContent = Get-Content -Path $SFTPLogFile -Raw
        
        # Add SFTP log to our main log file
        Write-Log "========== WINSCP LOG =========="
        Write-Log $logContent
        
        # Check for successful transfers
        foreach ($FileInfo in $FilesWithSupplierInfo) {
            $fileName   = $FileInfo.File.Name
            $fileStatus = "UNKNOWN"
            
            # Look for patterns of success in the log file
            if ($logContent -match "Transfer done: '.*$([regex]::Escape($fileName))' => .*" -or
                $logContent -match "$([regex]::Escape($fileName)).*\|.*\|.*\|.*\| 100%" -or
                $logContent -match "Upload of file '.*$([regex]::Escape($fileName)).*' was successful") {
                $fileStatus = "SUCCESS"
                $uploadSuccess = $true
            }
            elseif ($logContent -match "Error transferring file '.*$([regex]::Escape($fileName))'") {
                $fileStatus = "FAILED"
            }
            
            $resultMessage = "- $fileName - $fileStatus (Supplier: $($FileInfo.SupplierName), Code: $($FileInfo.SupplierCode))"
            $uploadResults += [PSCustomObject]@{
                FileName = $fileName
                FilePath = $FileInfo.File.FullName
                Status = $fileStatus
                SupplierCode = $FileInfo.SupplierCode
                SupplierName = $FileInfo.SupplierName
                SourceFile = $FileInfo.SourceFile
            }
            Write-Log $resultMessage
        }
    }
    catch {
        $logErrorMessage = "Error processing WinSCP log: $($_.Exception.Message)"
        Write-Log $logErrorMessage
    }
}
else {
    $logMissingMessage = "WinSCP log file not found at $SFTPLogFile. Unable to verify file transfer status."
    Write-Log $logMissingMessage
    
    foreach ($FileInfo in $FilesWithSupplierInfo) {
        $resultMessage = "- $($FileInfo.File.Name) - UNKNOWN (Log missing)"
        $uploadResults += [PSCustomObject]@{
            FileName = $FileInfo.File.Name
            FilePath = $FileInfo.File.FullName
            Status = "UNKNOWN (Log missing)"
            SupplierCode = $FileInfo.SupplierCode
            SupplierName = $FileInfo.SupplierName
            SourceFile = $FileInfo.SourceFile
        }
        Write-Log $resultMessage
    }
}

# Determine overall status
if ($exitCode -eq 0) {
    $statusMessage = "Files were uploaded successfully to SFTP server (Exit code: 0)."
    $uploadSuccess = $true
}
elseif ($uploadSuccess) {
    $statusMessage = "Some or all files were uploaded successfully, but WinSCP reported issues (Exit code: $exitCode)."
    $uploadSuccess = $true  # Consider it a success if some files transferred.
}
else {
    $statusMessage = "File upload FAILED. WinSCP exit code: $exitCode"
}

Write-Log $statusMessage

# Prepare for email notifications – copy files to a separate location for attachment purposes
$TempDir = Join-Path $env:TEMP "SFTPUpload_$Timestamp"
if (!(Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

$FileCopies = @()
foreach ($result in $uploadResults) {
    $TempPath = Join-Path $TempDir $result.FileName
    Copy-Item -Path $result.FilePath -Destination $TempPath -Force
    $FileCopies += $TempPath
}

$WScriptPath = "$Env:SystemRoot\System32\cscript.exe"

# Send individual detailed HTML emails for each file
foreach ($result in $uploadResults) {
    $fileName = $result.FileName
    $fileStatus = $result.Status
    $supplierCode = $result.SupplierCode
    $supplierName = $result.SupplierName
    $sourceFile = $result.SourceFile
    
    # Determine CSS class for status
    $statusClass = switch ($fileStatus) {
        "SUCCESS" { "success" }
        "FAILED" { "failed" }
        default { "unknown" }
    }
    
    # Find the temporary copy of the file
    $tempFilePath = Join-Path $TempDir $fileName
    
    # Construct the HTML email body for this file
    $HTMLBody = $HTMLHeader + @"
<h2>SFTP Upload Report</h2>
<div class="supplier-info">
    <p><strong>File:</strong> $fileName</p>
    <p><strong>Supplier Name:</strong> $supplierName</p>
    <p><strong>Supplier Code:</strong> $supplierCode</p>
    <p><strong>Status:</strong> <span class="$statusClass">$fileStatus</span></p>
    <p><strong>Source Supplier File:</strong> $sourceFile</p>
</div>
<p>$statusMessage</p>
<h3>Upload Results</h3>
<table>
  <tr>
    <th>File Name</th>
    <th>Supplier</th>
    <th>Code</th>
    <th>Status</th>
  </tr>
"@
    foreach ($res in $uploadResults) {
        $resStatusClass = switch ($res.Status) {
            "SUCCESS" { "success" }
            "FAILED" { "failed" }
            default { "unknown" }
        }
        $HTMLBody += "  <tr><td>$($res.FileName)</td><td>$($res.SupplierName)</td><td>$($res.SupplierCode)</td><td class='$resStatusClass'>$($res.Status)</td></tr>`r`n"
    }
    $HTMLBody += @"
</table>
<p>Detailed log can be found at: <em>$LogFile</em></p>
"@ + $HTMLFooter
    
    try {
        # Call the HTML email VBScript with the file as an attachment
        $emailOutput = & $WScriptPath $VBScriptPath "$EmailSubject - $fileName - $supplierName" "$tempFilePath" "$HTMLBody" 2>&1
        Write-Log "Email notification sent for file: $fileName (Supplier: $supplierName)"
        if ($emailOutput) {
            Write-Log "Email output: $emailOutput"
        }
        Start-Sleep -Seconds 1  # Small delay between emails
    }
    catch {
        $emailErrorMsg = "Error sending email notification for file $fileName. Error: $($_.Exception.Message)"
        Write-Log $emailErrorMsg
    }
}

# Also send a summary email with the main log file attached
$HTMLSummary = $HTMLHeader + @"
<h2>SFTP Upload Summary</h2>
<p>$statusMessage</p>
<h3>Files Processed</h3>
<table>
  <tr>
    <th>File Name</th>
    <th>Supplier</th>
    <th>Code</th>
    <th>Status</th>
  </tr>
"@
foreach ($res in $uploadResults) {
    $resStatusClass = switch ($res.Status) {
        "SUCCESS" { "success" }
        "FAILED" { "failed" }
        default { "unknown" }
    }
    $HTMLSummary += "  <tr><td>$($res.FileName)</td><td>$($res.SupplierName)</td><td>$($res.SupplierCode)</td><td class='$resStatusClass'>$($res.Status)</td></tr>`r`n"
}
$HTMLSummary += @"
</table>
<p>For more details, please refer to the attached log file.</p>
"@ + $HTMLFooter

try {
    $emailOutput = & $WScriptPath $VBScriptPath "$EmailSubject - Summary" "$LogFile" "$HTMLSummary" 2>&1
    Write-Log "Summary email notification sent with log file attachment."
    if ($emailOutput) {
        Write-Log "Email output: $emailOutput"
    }
}
catch {
    $summaryErrorMsg = "Error sending summary email notification. Error: $($_.Exception.Message)"
    Write-Log $summaryErrorMsg
}

# Archive files if upload was successful
if ($uploadSuccess) {
    Write-Log "Archiving files to $ArchiveDirectory..."
    
    foreach ($result in $uploadResults) {
        $DestinationPath = Join-Path -Path $ArchiveDirectory -ChildPath $result.FileName
        try {
            Move-Item -Path $result.FilePath -Destination $DestinationPath -Force
            Write-Log "- Archived $($result.FileName) to $ArchiveDirectory (Supplier: $($result.SupplierName))"
        }
        catch {
            Write-Log "- Failed to archive $($result.FileName). Error: $($_.Exception.Message)"
        }
    }
}
else {
    Write-Log "Files were not archived because upload could not be verified as successful."
}

# Clean up temporary files
if (Test-Path $ScriptPath) {
    try {
        Remove-Item -Path $ScriptPath -Force
        Write-Log "Temporary WinSCP script removed."
    }
    catch {
        Write-Log "Failed to remove temporary script file. Error: $($_.Exception.Message)"
    }
}

# Clean up temporary file copies
if (Test-Path $TempDir) {
    try {
        Remove-Item -Path $TempDir -Recurse -Force
        Write-Log "Temporary file copies removed."
    }
    catch {
        Write-Log "Failed to remove temporary file copies. Error: $($_.Exception.Message)"
    }
}

# Final log message
$completionMessage = "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log $completionMessage

# Return appropriate exit code
if ($uploadSuccess) {
    exit 0
}
else {
    exit 1
}
