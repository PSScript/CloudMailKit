<#
.SYNOPSIS
    Test harness for CloudMailReader COM DLL
.DESCRIPTION
    Validates all CloudMailReader functions
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$true)]
    [string]$MailboxAddress
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CloudMailReader Test Harness" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Create reader instance
    Write-Host "[1/8] Creating GraphMailClient..." -ForegroundColor Yellow
    $reader = New-Object -ComObject CloudMailReader.GraphMailClient
    Write-Host "      ✓ Success" -ForegroundColor Green
    Write-Host ""

    # Initialize
    Write-Host "[2/8] Initializing with credentials..." -ForegroundColor Yellow
    $reader.Initialize($TenantId, $ClientId, $ClientSecret, $MailboxAddress)
    Write-Host "      ✓ Success (token acquired)" -ForegroundColor Green
    Write-Host ""

    # Get inbox
    Write-Host "[3/8] Getting Inbox folder..." -ForegroundColor Yellow
    $inbox = $reader.GetInbox()
    Write-Host "      ✓ Folder: $($inbox.DisplayName)" -ForegroundColor Green
    Write-Host "      ✓ Total: $($inbox.TotalItemCount) messages" -ForegroundColor Green
    Write-Host "      ✓ Unread: $($inbox.UnreadItemCount) messages" -ForegroundColor Green
    Write-Host ""

    # List folders
    Write-Host "[4/8] Listing all folders..." -ForegroundColor Yellow
    $folders = $reader.ListFolders()
    Write-Host "      ✓ Found $($folders.Count) folders:" -ForegroundColor Green
    foreach ($folder in $folders) {
        Write-Host "        - $($folder.DisplayName) ($($folder.TotalItemCount) items)" -ForegroundColor Gray
    }
    Write-Host ""

    # List messages
    Write-Host "[5/8] Listing messages from Inbox..." -ForegroundColor Yellow
    $messages = $reader.ListMessages($inbox.Id, 10)
    Write-Host "      ✓ Retrieved $($messages.Count) messages" -ForegroundColor Green
    
    if ($messages.Count -gt 0) {
        Write-Host "      ✓ Sample messages:" -ForegroundColor Green
        foreach ($msg in $messages | Select-Object -First 3) {
            Write-Host "        - $($msg.Subject)" -ForegroundColor Gray
            Write-Host "          From: $($msg.FromName) <$($msg.FromAddress)>" -ForegroundColor DarkGray
            Write-Host "          Date: $($msg.ReceivedDateTime)" -ForegroundColor DarkGray
        }
    }
    Write-Host ""

    # Get MIME
    if ($messages.Count -gt 0) {
        $firstMsg = $messages[0]
        Write-Host "[6/8] Fetching raw MIME for first message..." -ForegroundColor Yellow
        $mime = $reader.GetMessageMime($firstMsg.Id)
        $mimeLength = $mime.Length
        Write-Host "      ✓ MIME fetched ($mimeLength bytes)" -ForegroundColor Green
        Write-Host ""

        # Parse MIME
        Write-Host "[7/8] Parsing MIME with MimeParser..." -ForegroundColor Yellow
        $parser = New-Object -ComObject CloudMailReader.MimeParser
        
        $subject = $parser.GetSubject($mime)
        $from = $parser.GetFromAddress($mime)
        $textBody = $parser.GetTextBody($mime)
        $attachCount = $parser.GetAttachmentCount($mime)
        
        Write-Host "      ✓ Subject: $subject" -ForegroundColor Green
        Write-Host "      ✓ From: $from" -ForegroundColor Green
        Write-Host "      ✓ Body length: $($textBody.Length) chars" -ForegroundColor Green
        Write-Host "      ✓ Attachments: $attachCount" -ForegroundColor Green
        Write-Host ""

        # Test operations
        Write-Host "[8/8] Testing message operations..." -ForegroundColor Yellow
        
        if (-not $firstMsg.IsRead) {
            $reader.MarkAsRead($firstMsg.Id)
            Write-Host "      ✓ Marked message as read" -ForegroundColor Green
            Start-Sleep -Seconds 1
            $reader.MarkAsUnread($firstMsg.Id)
            Write-Host "      ✓ Marked message as unread" -ForegroundColor Green
        }
        Write-Host ""
    } else {
        Write-Host "[6/8] Skipped - No messages to test" -ForegroundColor DarkYellow
        Write-Host "[7/8] Skipped - No messages to test" -ForegroundColor DarkYellow
        Write-Host "[8/8] Skipped - No messages to test" -ForegroundColor DarkYellow
        Write-Host ""
    }

    Write-Host "========================================" -ForegroundColor Green
    Write-Host "ALL TESTS PASSED" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "TEST FAILED" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor DarkRed
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}