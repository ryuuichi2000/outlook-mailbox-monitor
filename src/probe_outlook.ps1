[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string] $MailboxSmtp,
    [Parameter(Mandatory = $true)] [int] $RetentionHours,
    [string] $OutFile = ''
)

# READ-ONLY preflight + inbox census. Writes UTF-8 JSON to a file so the child
# process stdout encoding never affects parsing.
# No item is modified, moved, or deleted anywhere in this script.
# All environment-specific values come from parameters (see config.json).
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($OutFile)) {
    $OutFile = [IO.Path]::Combine($PSScriptRoot, 'probe_result.json')
}
if ($RetentionHours -le 0 -or [string]::IsNullOrWhiteSpace($MailboxSmtp)) {
    throw 'InvalidConfiguration'
}
$utf8 = [Text.UTF8Encoding]::new($false)
$DeliveryTag = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'

$r = [ordered] @{
    ok = $false
    reason = ''
    checked_utc = [DateTimeOffset]::UtcNow.ToString('o')
    outlook_version = ''
    accounts = @()
    target = $null
    inbox = $null
    subfolders = @()
    error_code = ''
}
function Save-Result {
    [IO.File]::WriteAllText($OutFile, ($r | ConvertTo-Json -Compress -Depth 8), $utf8)
}

try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $r.outlook_version = [string] $app.Version
    $session = $app.Session
    $accts = $session.Accounts
    $account = $null
    for ($i = 1; $i -le $accts.Count; $i++) {
        $a = $accts.Item($i)
        $r.accounts += [string] $a.SmtpAddress
        if ($a.SmtpAddress -eq $MailboxSmtp) { $account = $a }
    }
    if ($null -eq $account) { throw 'target_account_not_configured' }

    $store = $account.DeliveryStore
    $r.target = [ordered] @{
        smtp = [string] $account.SmtpAddress
        connection_mode = [int] $account.ExchangeConnectionMode
        store_type = [int] $store.ExchangeStoreType
        is_data_file_store = [bool] $store.IsDataFileStore
    }

    $inbox = $store.GetDefaultFolder(6)
    $items = $inbox.Items
    $cutoff = [DateTimeOffset]::UtcNow.AddHours(-$RetentionHours).UtcDateTime
    $census = [ordered] @{
        total_items = [int] $items.Count
        non_mail_items = 0
        missing_delivery_time = 0
        older_than_cutoff = 0
        cutoff_utc = ([DateTimeOffset] $cutoff).ToString('o')
        oldest_received = ''
        newest_received = ''
    }
    for ($i = 1; $i -le $items.Count; $i++) {
        $it = $items.Item($i)
        if ($it.Class -ne 43) { $census.non_mail_items++; continue }
        $d = $null
        try { $d = $it.PropertyAccessor.GetProperty($DeliveryTag) } catch { }
        if ($d -isnot [datetime]) { $census.missing_delivery_time++; continue }
        $du = [datetime]::SpecifyKind($d, [DateTimeKind]::Utc)
        if ($du -lt $cutoff) { $census.older_than_cutoff++ }
    }
    # ReceivedTime extremes (read-only; Sort affects only this local view).
    try {
        $items.Sort('[ReceivedTime]', $false)
        if ($items.Count -gt 0) { $census.oldest_received = $items.Item(1).ReceivedTime.ToString('o') }
        $items.Sort('[ReceivedTime]', $true)
        if ($items.Count -gt 0) { $census.newest_received = $items.Item(1).ReceivedTime.ToString('o') }
    } catch { }
    $r.inbox = $census

    try {
        $subs = $inbox.Folders
        for ($j = 1; $j -le $subs.Count; $j++) {
            $f = $subs.Item($j)
            $r.subfolders += [ordered] @{
                name = [string] $f.Name
                items = [int] $f.Items.Count
                subfolder_count = [int] $f.Folders.Count
            }
        }
    } catch { }

    $r.ok = $true
    $r.reason = 'outlook_ready'
    Save-Result
} catch {
    $r.reason = [string] $_.Exception.GetBaseException().Message
    $r.error_code = '0x{0:X8}' -f $_.Exception.GetBaseException().HResult
    Save-Result
}
