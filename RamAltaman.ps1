#Requires -Version 5.1

<#
.SYNOPSIS
    Reports unused RAM to Sam Altman, at most once per day.

.EXAMPLE
    .\ram-altman.ps1 -Status

.EXAMPLE
    .\ram-altman.ps1 -WhatIf

.EXAMPLE
    .\ram-altman.ps1 -ThresholdPercent 25 -WhatIf -Verbose
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param (
    [switch]$Status,

    [string]$Recipient = $env:RAM_ALTMAN_TO,
    [string]$Sender = $env:RAM_ALTMAN_FROM,
    [string]$SenderName = $env:RAM_ALTMAN_FROM_NAME,

    [ValidateRange(0, 100)]
    [double]$ThresholdPercent = $(
        if ($env:RAM_ALTMAN_THRESHOLD) {
            $env:RAM_ALTMAN_THRESHOLD
        }
        else {
            10
        }
    ),

    [string]$StateDirectory = $(
        if ($env:RAM_ALTMAN_STATE_DIR) {
            $env:RAM_ALTMAN_STATE_DIR
        }
        else {
            Join-Path $env:LOCALAPPDATA 'ram-altman'
        }
    ),

    [string]$SmtpServer = $env:RAM_ALTMAN_SMTP_SERVER,

    [ValidateRange(1, 65535)]
    [int]$SmtpPort = $(
        if ($env:RAM_ALTMAN_SMTP_PORT) {
            $env:RAM_ALTMAN_SMTP_PORT
        }
        else {
            25
        }
    ),

    [switch]$UseSsl,
    [pscredential]$Credential,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$UsageLog    = Join-Path $StateDirectory 'usage.csv'
$LastSent    = Join-Path $StateDirectory 'last_sent'
$LockFile    = Join-Path $StateDirectory 'lock'
$MaxLogBytes = 10MB
$LogLines    = 50000

function Get-MemoryInfo {

    $OS = Get-CimInstance Win32_OperatingSystem

    $Total     = [long]$OS.TotalVisibleMemorySize
    $Available = [long]$OS.FreePhysicalMemory
    $Used      = $Total - $Available
    $Percent   = [math]::Round(($Available * 100.0) / $Total, 1)

    [pscustomobject]@{
        TotalKB          = $Total
        UsedKB           = $Used
        AvailableKB      = $Available
        AvailablePercent = $Percent
    }
}

function ConvertTo-GiB ([long]$KB) {
    '{0:N1} GiB' -f ($KB / 1MB)
}

function Show-MemoryStatus ($Memory) {

    Write-Host ('Total:      {0}' -f (ConvertTo-GiB $Memory.TotalKB))
    Write-Host ('Used:       {0}' -f (ConvertTo-GiB $Memory.UsedKB))
    Write-Host ('Available:  {0} ({1:N1}%)' -f
        (ConvertTo-GiB $Memory.AvailableKB),
        $Memory.AvailablePercent)
}

function Write-MemorySample ($Memory) {

    if (-not $PSCmdlet.ShouldProcess($UsageLog, 'Append memory sample')) {
        return
    }

    if (-not (Test-Path $UsageLog)) {
        'timestamp,total_kb,used_kb,available_kb,available_pct' |
            Set-Content $UsageLog -Encoding UTF8
    }

    $Time = Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz'

    '{0},{1},{2},{3},{4:N1}' -f
        $Time,
        $Memory.TotalKB,
        $Memory.UsedKB,
        $Memory.AvailableKB,
        $Memory.AvailablePercent |
        Add-Content $UsageLog -Encoding UTF8

    if ((Get-Item $UsageLog).Length -gt $MaxLogBytes) {

        $Temp = "$UsageLog.tmp"

        $Header = Get-Content $UsageLog -First 1
        $Tail   = Get-Content $UsageLog -Tail $LogLines

        @($Header) + @($Tail) |
            Set-Content $Temp -Encoding UTF8

        Move-Item $Temp $UsageLog -Force
    }
}

function Get-TemplateIndex {
    param (
        [string]$Name,
        [int]$Count
    )

    # -WhatIf must not modify the shuffle bag.
    if ($WhatIfPreference) {
        return Get-Random -Minimum 0 -Maximum $Count
    }

    $Bag      = Join-Path $StateDirectory "bag_$Name"
    $LastFile = "$Bag.last"

    $Indexes = @()

    if (Test-Path $Bag) {
        $Indexes = @(
            Get-Content $Bag |
                Where-Object { $_ -match '^\d+$' } |
                ForEach-Object { [int]$_ }
        )
    }

    if (-not $Indexes.Count) {

        $Indexes = @(0..($Count - 1) | Sort-Object { Get-Random })

        if (Test-Path $LastFile) {

            $Previous = [int](Get-Content $LastFile -Raw)

            if ($Count -gt 1 -and $Indexes[0] -eq $Previous) {
                $Indexes = @($Indexes[1..($Count - 1)] + $Indexes[0])
            }
        }
    }

    $Next = $Indexes[0]

    if ($Indexes.Count -gt 1) {
        $Indexes[1..($Indexes.Count - 1)] |
            Set-Content $Bag
    }
    else {
        Clear-Content $Bag
    }

    $Next | Set-Content $LastFile

    if ($Next -ge $Count) {
        $Next = Get-Random -Minimum 0 -Maximum $Count
    }

    return $Next
}

function New-RamMessage ($Memory) {

    $Total     = ConvertTo-GiB $Memory.TotalKB
    $Available = ConvertTo-GiB $Memory.AvailableKB
    $Pct       = '{0:N1}' -f $Memory.AvailablePercent

    $Subjects = @(
        'Good news: we found some RAM'
        'Unused RAM detected in the wild'
        'RAM sighting on my PC'
        "$Available of RAM, just sitting there"
        'You missed a spot'
        'Daily RAM census: some survived'
        "$Pct% of my RAM remains at large"
        "Still not all of it, I'm afraid"
    )

    $Bodies = @(
        "I'm writing to report that you appear to have missed some. This PC has $Total of RAM and $Available of it ($Pct%) is currently available. I assume this is an oversight."

        "A routine inspection of my PC found $Available of unclaimed RAM out of $Total. That's $Pct% of the machine doing absolutely nothing. I thought you'd want to know before someone tries to train something on it."

        @"
Daily status report:

  Total RAM:     $Total
  Available:     $Available ($Pct%)
  Threat level:  someone at OpenAI might notice
"@

        "Even with every browser tab I own open, $Available of my $Total of RAM is still available. I'm as surprised as you are."

        "Please don't be alarmed, but there is $Available of RAM on my desk that isn't in a data centre. That's $Pct% of what's installed here, and as far as I can tell it has no plans today."

        "Today's forecast for my PC: clear skies, light workloads and $Pct% available memory. Expect $Available of the installed $Total to stay idle through the evening."

        "I hear RAM is hard to come by these days, so it feels a bit rude to be sitting on $Available of it. Total installed: $Total. It's not for sale, I just like knowing it's there."

        "I have a confession. I have $Available of RAM that I'm not using, which is $Pct% of the $Total in this PC. I've told nobody but you and a scheduled task."

        "If $Available of RAM sits idle in a PC and nobody allocates it, does it still count as capacity? Asking on behalf of the $Pct% of my memory with nothing to do."
    )

    $Signoffs = @(
        'Yours in abundance,'
        'Regards from the last free RAM in the county,'
        'Hoarding responsibly,'
        'Stay allocated,'
        'Yours, with memory to spare,'
        'Kind regards,'
    )

    $Subject = $Subjects[
        (Get-TemplateIndex subject $Subjects.Count)
    ]

    $Body = $Bodies[
        (Get-TemplateIndex body $Bodies.Count)
    ]

    $Signoff = $Signoffs[
        (Get-TemplateIndex signoff $Signoffs.Count)
    ]

    if (-not $SenderName) {
        $SenderName = 'A concerned RAM owner'
    }

    [pscustomobject]@{
        Subject = $Subject

        Body = @"
Hi Sam,

$Body

$Signoff
$SenderName

--
Sent automatically, at most once a day, by a small script on one PC.
Reply if you'd like it to stop.
"@
    }
}

function Show-Message ($Message) {

    Write-Host ''
    Write-Host "From: $SenderName <$Sender>"
    Write-Host "To: $Recipient"
    Write-Host "Subject: $($Message.Subject)"
    Write-Host ''
    Write-Host $Message.Body
}

function Send-RamMessage ($Message) {

    if (-not $PSCmdlet.ShouldProcess(
        $Recipient,
        "Send email '$($Message.Subject)'"
    )) {
        Show-Message $Message
        return $false
    }

    $Mail = New-Object System.Net.Mail.MailMessage

    if ($SenderName) {
        $Mail.From = New-Object System.Net.Mail.MailAddress(
            $Sender,
            $SenderName
        )
    }
    else {
        $Mail.From = $Sender
    }

    $Mail.To.Add($Recipient)
    $Mail.Subject = $Message.Subject
    $Mail.Body = $Message.Body
    $Mail.BodyEncoding = [Text.Encoding]::UTF8
    $Mail.Headers.Add('Auto-Submitted', 'auto-generated')

    $SMTP = New-Object System.Net.Mail.SmtpClient(
        $SmtpServer,
        $SmtpPort
    )

    try {
        $SMTP.EnableSsl = $UseSsl.IsPresent

        if ($Credential) {
            $SMTP.UseDefaultCredentials = $false
            $SMTP.Credentials = $Credential.GetNetworkCredential()
        }

        $SMTP.Send($Mail)
    }
    finally {
        $Mail.Dispose()
        $SMTP.Dispose()
    }

    return $true
}


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

try {

    $Memory = Get-MemoryInfo

    if ($Status) {
        Show-MemoryStatus $Memory
        return
    }

    # During -WhatIf, fake addresses are sufficient for the preview.
    if ($WhatIfPreference) {
        if (-not $Recipient) { $Recipient = 'recipient-not-set@example.com' }
        if (-not $Sender)    { $Sender    = 'sender-not-set@example.com' }
    }
    else {
        if (-not $Recipient) {
            throw 'Recipient is not set.'
        }

        if (-not $Sender) {
            throw 'Sender is not set.'
        }

        if (-not $SmtpServer) {
            throw 'SMTP server is not set.'
        }
    }

    if (-not (Test-Path $StateDirectory)) {
        if ($PSCmdlet.ShouldProcess($StateDirectory, 'Create state directory')) {
            $null = New-Item $StateDirectory -ItemType Directory -Force
        }
    }

    # -WhatIf deliberately does not create/open the persistent lock.
    $Lock = $null

    if (-not $WhatIfPreference) {
        try {
            $Lock = [IO.File]::Open(
                $LockFile,
                [IO.FileMode]::OpenOrCreate,
                [IO.FileAccess]::ReadWrite,
                [IO.FileShare]::None
            )
        }
        catch [IO.IOException] {
            Write-Verbose 'Another instance is already running.'
            return
        }
    }

    try {

        Write-MemorySample $Memory

        if ($Memory.AvailablePercent -le $ThresholdPercent) {
            Write-Verbose (
                'Available RAM ({0:N1}%) does not exceed threshold ({1:N1}%).' -f
                $Memory.AvailablePercent,
                $ThresholdPercent
            )

            return
        }

        $Today = Get-Date -Format 'yyyy-MM-dd'

        if ((Test-Path $LastSent) -and -not $Force) {

            if ((Get-Content $LastSent -Raw).Trim() -eq $Today) {
                Write-Verbose 'The daily message has already been sent.'
                return
            }
        }

        $Message = New-RamMessage $Memory
        $Sent = Send-RamMessage $Message

        if ($Sent -and $PSCmdlet.ShouldProcess(
            $LastSent,
            "Record successful send for $Today"
        )) {
           $Today | Set-Content $LastSent
        }
    }
    finally {
        if ($Lock) {
            $Lock.Dispose()
        }
    }
}
catch {
    Write-Error "ram-altman: $($_.Exception.Message)"
    exit 1
}