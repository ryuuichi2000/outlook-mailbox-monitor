[CmdletBinding()]
param(
    [ValidateSet('Check', 'Preview', 'Archive', 'ResolveMove')]
    [string] $Mode = 'Archive',

    # Every environment-specific value is supplied by the caller from
    # config.json. No mailbox address, quota or path is embedded in this file.
    [Parameter(Mandatory = $true)] [string] $MailboxSmtp,
    [Parameter(Mandatory = $true)] [long] $QuotaBytes,
    [Parameter(Mandatory = $true)] [double] $TriggerRatio,
    [Parameter(Mandatory = $true)] [string] $DataDirectory,
    [Parameter(Mandatory = $true)] [string] $ArchiveFolderName,
    [Parameter(Mandatory = $true)] [string] $PstFilePrefix,
    [Parameter(Mandatory = $true)] [string] $TimeZoneId,
    [Parameter(Mandatory = $true)] [int] $RetentionHours,
    [long] $DiskReserveBytes = 134217728,
    [long] $PstCeilingBytes = 48318382080,

    [switch] $AllowCachedArchive,

    [ValidatePattern('^[a-f0-9]{32}$')]
    [string] $RunId,

    [ValidateSet('Archived', 'StillInInbox')]
    [string] $MoveResolution,

    [string] $ArchivedEntryId,

    [ValidateRange(1, 500)]
    [int] $MaxMessagesPerRun = 100,

    # Daily-once mode: read usage, archive, settle and publish all happen inside
    # a single invocation. No cross-run Pending/InFlight state is read or written,
    # and the retention cutoff is applied to ALL eligible messages without a cap.
    [switch] $DailyOnce
)

# Generic, configuration-driven build. See README.md and docs/SETUP.md.
# Run in the interactive Windows session that already has classic Outlook open.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if ($QuotaBytes -le 0 -or $TriggerRatio -le 0 -or $TriggerRatio -gt 1 -or
    $RetentionHours -le 0 -or [string]::IsNullOrWhiteSpace($MailboxSmtp) -or
    [string]::IsNullOrWhiteSpace($DataDirectory)) { throw 'InvalidConfiguration' }

$Mailbox = $MailboxSmtp
$TriggerBytes = [long] [Math]::Ceiling($QuotaBytes * $TriggerRatio)
$ArchiveTimeZone = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
$VerificationDelay = [TimeSpan]::FromMinutes(5)
$VerificationTimeout = [TimeSpan]::FromMinutes(30)
$Utf8 = [Text.UTF8Encoding]::new($false)
$script:ExitCode = 0

function Get-DatedPstPath {
    param([DateTimeOffset] $StartedUtc)

    $date = [TimeZoneInfo]::ConvertTime($StartedUtc, $ArchiveTimeZone)
    $suffix = $date.ToString('yyMMdd', [Globalization.CultureInfo]::InvariantCulture)
    return [IO.Path]::Combine($DataDirectory, "$PstFilePrefix-$suffix.pst")
}

function Test-RegularPath {
    param([string] $Path, [switch] $Directory)

    try {
        $attributes = [IO.File]::GetAttributes($Path)
    } catch {
        $cause = $_.Exception.GetBaseException()
        if ($cause -is [IO.FileNotFoundException] -or $cause -is [IO.DirectoryNotFoundException]) { return $false }
        throw
    }
    if ($attributes -band [IO.FileAttributes]::ReparsePoint) { throw "ReparsePointRejected: $Path" }
    if ([bool] ($attributes -band [IO.FileAttributes]::Directory) -ne [bool] $Directory) {
        throw "UnexpectedPathType: $Path"
    }
    return $true
}

function Initialize-DataDirectory {
    $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($DataDirectory))
    if (-not $drive.IsReady -or $drive.DriveType -ne [IO.DriveType]::Fixed -or
        $drive.DriveFormat -notin @('NTFS', 'ReFS')) {
        throw 'ArchiveDriveUnavailable: a fixed local NTFS or ReFS D: drive is required.'
    }

    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $trustedSids = @($currentUser, 'S-1-5-18', 'S-1-5-32-544') # user, SYSTEM, Administrators
    if (-not (Test-RegularPath -Path $DataDirectory -Directory)) {
        $security = [Security.AccessControl.DirectorySecurity]::new()
        $security.SetAccessRuleProtection($true, $false)
        foreach ($sid in $trustedSids) {
            $identity = [Security.Principal.SecurityIdentifier]::new($sid)
            $rule = [Security.AccessControl.FileSystemAccessRule]::new($identity, 'FullControl',
                'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $security.AddAccessRule($rule)
        }
        [void] [IO.Directory]::CreateDirectory($DataDirectory, $security)
    }
    if (-not (Test-RegularPath -Path $DataDirectory -Directory)) { throw 'OutputDirectoryUnavailable' }
    Assert-PrivateAcl -Security ([IO.Directory]::GetAccessControl($DataDirectory)) -Path $DataDirectory

    $eventDirectory = Join-Path $DataDirectory 'events'
    if (Test-RegularPath -Path $eventDirectory -Directory) {
        Assert-PrivateAcl -Security ([IO.Directory]::GetAccessControl($eventDirectory)) -Path $eventDirectory
    }
}

function Assert-PrivateAcl {
    param($Security, [string] $Path)

    $sidType = [Security.Principal.SecurityIdentifier]
    $trusted = @([Security.Principal.WindowsIdentity]::GetCurrent().User.Value, 'S-1-5-18', 'S-1-5-32-544')
    if ($Security.GetOwner($sidType).Value -notin $trusted) { throw "UntrustedOutputOwner: $Path" }
    foreach ($rule in $Security.GetAccessRules($true, $true, $sidType)) {
        if ($rule.AccessControlType -eq 'Allow' -and $rule.IdentityReference.Value -notin $trusted) {
            throw "UntrustedOutputAcl: $Path"
        }
    }
}

function Test-PrivateFile {
    param([string] $Path)

    if (-not (Test-RegularPath -Path $Path)) { return $false }
    Assert-PrivateAcl -Security ([IO.File]::GetAccessControl($Path)) -Path $Path
    return $true
}

function Write-Log {
    param([string] $Name, [System.Collections.IDictionary] $Details = @{})

    $record = [ordered] @{
        time_utc = [DateTimeOffset]::UtcNow.ToString('o')
        mailbox = $Mailbox
        mode = $Mode
        action = $Name
        details = $Details
    }
    $json = $record | ConvertTo-Json -Compress -Depth 8
    $logPath = Join-Path $DataDirectory ('monitor-{0:yyyy-MM}.jsonl' -f [datetime]::UtcNow)
    [void] (Test-PrivateFile -Path $logPath)
    $stream = [IO.File]::Open($logPath, 'OpenOrCreate', 'ReadWrite', 'Read')
    try {
        if ($stream.Length -gt 0) {
            [void] $stream.Seek(-1, [IO.SeekOrigin]::End)
            if ($stream.ReadByte() -ne 10) { $stream.WriteByte(10) }
        }
        $bytes = $Utf8.GetBytes($json + "`n")
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()
    } finally {
        $stream.Dispose()
    }
    Write-Output $json
}

function Write-JsonFile {
    param([string] $Path, $Value)

    $temporaryPath = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    $stream = [IO.File]::Open($temporaryPath, 'CreateNew', 'Write', 'None')
    try {
        $bytes = $Utf8.GetBytes(($Value | ConvertTo-Json -Depth 8))
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally {
        $stream.Dispose()
    }
    if (Test-PrivateFile -Path $Path) {
        [IO.File]::Replace($temporaryPath, $Path, [NullString]::Value)
    } else {
        [IO.File]::Move($temporaryPath, $Path)
    }
}

function Save-State {
    Write-JsonFile -Path (Join-Path $DataDirectory 'state.json') -Value $script:State
}

function Read-State {
    $path = Join-Path $DataDirectory 'state.json'
    if (-not (Test-PrivateFile -Path $path)) {
        return [pscustomobject] @{ Version = 2; Pending = $null; IssueReason = ''; IssueId = ''; UnavailableCount = 0 }
    }
    if ([IO.FileInfo]::new($path).Length -gt 64KB) { throw 'StateInvalid' }
    $saved = [IO.File]::ReadAllText($path) | ConvertFrom-Json
    if ($saved.Version -ne 2 -or $saved.IssueReason -isnot [string] -or
        $saved.IssueId -notmatch '^(?:[a-f0-9]{32})?$' -or $saved.UnavailableCount -lt 0) {
        throw 'StateInvalid'
    }
    if ($null -ne $saved.Pending) {
        $pending = $saved.Pending
        $cutoff = [DateTimeOffset]::Parse($pending.CutoffUtc)
        $datedPath = Get-DatedPstPath -StartedUtc ($cutoff.AddHours($RetentionHours))
        # Allow an unfinished cycle created by the earlier fixed-name release to resume.
        $legacyPath = [IO.Path]::Combine($DataDirectory, "$PstFilePrefix.pst")
        if ($pending.Id -notmatch '^[a-f0-9]{32}$' -or $pending.PstPath -notin @($datedPath, $legacyPath) -or
            $cutoff -gt [DateTimeOffset]::UtcNow.AddHours(-$RetentionHours) -or $pending.MovedCount -lt 0 -or
            $pending.BeforeBytes -lt $TriggerBytes -or $pending.LastBatchBeforeBytes -lt 0) {
            throw 'StateInvalid'
        }
        if (($pending.MovedCount -gt 0 -or $null -ne $pending.InFlight) -and -not $pending.PstCreatedUtc) {
            throw 'StateInvalid'
        }
        if ($pending.MovedCount -gt 0 -and [DateTimeOffset]::Parse($pending.LastMoveUtc) -gt [DateTimeOffset]::UtcNow) {
            throw 'StateInvalid'
        }
        if ($null -ne $pending.InFlight -and
            (-not $pending.InFlight.EntryId -or -not $pending.InFlight.StoreId -or $pending.InFlight.AttachmentCount -lt 0)) {
            throw 'StateInvalid'
        }
    }
    return $saved
}

function Publish-Event {
    param([ValidatePattern('^(archive-)?[a-f0-9]{32}$')] [string] $Id,
        [string] $Type, [System.Collections.IDictionary] $Details)

    $eventDirectory = Join-Path $DataDirectory 'events'
    if (-not (Test-RegularPath -Path $eventDirectory -Directory)) { [void] [IO.Directory]::CreateDirectory($eventDirectory) }
    $eventPath = Join-Path $eventDirectory "$Id.json"
    if (-not (Test-PrivateFile -Path $eventPath)) {
        Write-JsonFile -Path $eventPath -Value ([ordered] @{
            schema_version = 1
            event_id = $Id
            event_type = $Type
            time_utc = [DateTimeOffset]::UtcNow.ToString('o')
            mailbox = $Mailbox
            quota_bytes = $QuotaBytes
            trigger_bytes = $TriggerBytes
            details = $Details
        })
    }
}

function Publish-Issue {
    param([string] $Reason, [System.Collections.IDictionary] $Details = @{})

    # Persist the ID before publishing. An interrupted write can retry the same event.
    if ($script:State.IssueReason -ne $Reason) {
        $script:State.IssueReason = $Reason
        $script:State.IssueId = [guid]::NewGuid().ToString('N')
        Save-State
    }
    $Details['reason'] = $Reason
    Publish-Event -Id $script:State.IssueId -Type 'attention_required' -Details $Details
}

function Get-ErrorCode {
    param([System.Management.Automation.ErrorRecord] $Record)

    return ('0x{0:X8}' -f $Record.Exception.GetBaseException().HResult)
}

function Get-MailboxContext {
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    $session = $outlook.Session
    $accounts = $session.Accounts
    $matchingAccount = $null
    for ($i = 1; $i -le $accounts.Count; $i++) {
        $account = $accounts.Item($i)
        if ($account.SmtpAddress -ne $Mailbox) { continue }
        if ($null -ne $matchingAccount) { throw 'AmbiguousMailbox' }
        $matchingAccount = $account
    }
    if ($null -eq $matchingAccount) { throw 'MailboxNotConfigured' }
    if ($matchingAccount.AccountType -ne 0) { throw 'MailboxIsNotExchange' }
    $store = $matchingAccount.DeliveryStore
    if ($null -eq $store -or $store.ExchangeStoreType -notin @(0, 1, 4)) {
        throw 'ExchangeDeliveryStoreUnavailable'
    }
    return [pscustomobject] @{ Session = $session; Account = $matchingAccount; Store = $store }
}

function Get-AccountConnection {
    param($Context)

    $modeNumber = [int] $Context.Account.ExchangeConnectionMode
    return [pscustomobject] @{
        ModeNumber = $modeNumber
        Connected = $modeNumber -in @(500, 600, 700, 800)
        ArchiveAllowed = $modeNumber -eq 800 -or ($AllowCachedArchive -and $modeNumber -eq 700)
        Source = if ($modeNumber -eq 800) { 'outlook_online' } else { 'outlook_cached_or_offline' }
    }
}

function Get-Usage {
    param($Context)

    $connection = Get-AccountConnection -Context $Context
    $accessor = $Context.Store.PropertyAccessor
    $used = $accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E080014')
    if (($used -isnot [long] -and $used -isnot [int]) -or $used -lt 0) { throw 'InvalidMailboxUsage' }

    $sendQuotaBytes = $null
    foreach ($tag in @('341B0003', '666E0003')) {
        try {
            $quotaKiB = [long] $accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x$tag")
            if ($quotaKiB -gt 0 -and $quotaKiB -lt [int]::MaxValue) {
                $sendQuotaBytes = $quotaKiB * 1KB
                break
            }
        } catch {
            # Some Outlook providers expose only one of the two properties.
        }
    }
    $currentConnection = Get-AccountConnection -Context $Context
    $connectionStable = $connection.ModeNumber -eq $currentConnection.ModeNumber
    return [pscustomobject] @{
        UsedBytes = [long] $used
        SendQuotaBytes = $sendQuotaBytes
        Connection = $currentConnection.ModeNumber
        Connected = $connectionStable -and $currentConnection.Connected
        ArchiveAllowed = $connectionStable -and $currentConnection.ArchiveAllowed
        Source = if ($connectionStable) { $currentConnection.Source } else { 'outlook_connection_transition' }
    }
}

function Assert-PstSpace {
    param([string] $Path, [long] $MessageBytes = 0)

    $drive = [IO.DriveInfo]::new([IO.Path]::GetPathRoot($Path))
    if ($drive.AvailableFreeSpace -lt ($DiskReserveBytes + $MessageBytes)) { throw 'PstDiskSpaceLow' }
    if (Test-PrivateFile -Path $Path) {
        $size = [IO.FileInfo]::new($Path).Length
        if ($size + $MessageBytes + $DiskReserveBytes -ge $PstCeilingBytes) { throw 'PstCapacityNearLimit' }
    }
}

function Assert-PendingPst {
    param($Pending)

    if ($Pending.PstCreatedUtc -and (-not (Test-PrivateFile -Path $Pending.PstPath) -or
            [IO.File]::GetCreationTimeUtc($Pending.PstPath).ToString('o') -ne $Pending.PstCreatedUtc)) {
        throw 'PendingPstMissingOrReplaced'
    }
}

function Find-PstStore {
    param($Session, [string] $Path)

    $stores = $Session.Stores
    for ($i = 1; $i -le $stores.Count; $i++) {
        $store = $stores.Item($i)
        if ($store.IsDataFileStore -and $store.FilePath -ieq $Path) {
            if ($store.ExchangeStoreType -ne 3) { throw 'The destination is not a local PST store.' }
            return $store
        }
    }
}

function Get-PstFolder {
    param($Context, [string] $Path)

    Assert-PstSpace -Path $Path
    $destinationStore = Find-PstStore -Session $Context.Session -Path $Path
    if ($null -eq $destinationStore) {
        [void] $Context.Session.AddStoreEx($Path, 2) # olStoreUnicode
        $destinationStore = Find-PstStore -Session $Context.Session -Path $Path
    }
    if ($null -eq $destinationStore -or $destinationStore.StoreID -eq $Context.Store.StoreID) {
        throw 'PstStoreUnavailable'
    }

    $root = $destinationStore.GetRootFolder()
    $folders = $root.Folders
    for ($i = 1; $i -le $folders.Count; $i++) {
        $folder = $folders.Item($i)
        if ($folder.Name -eq $ArchiveFolderName) { return $folder }
    }
    return $folders.Add($ArchiveFolderName, 6) # olFolderInbox, creates a mail folder
}

function Get-MailDeliveryUtc {
    param($Item)

    $value = $Item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E060040')
    if ($value -isnot [datetime] -or $value.Year -lt 1970 -or $value.Year -ge 4500) { throw 'InvalidDeliveryTime' }
    return [datetime]::SpecifyKind($value, [DateTimeKind]::Utc)
}

function Get-ArchiveCandidates {
    param($Inbox, [datetime] $CutoffUtc, [int] $Limit)

    $items = $Inbox.Items
    $items.Sort('[ReceivedTime]', $false)
    for ($i = 1; $i -le $items.Count -and $Limit -gt 0; $i++) {
        $item = $items.Item($i)
        if ($item.Class -ne 43) { continue } # olMail, excludes meeting requests and other classes
        if ((Get-MailDeliveryUtc -Item $item) -ge $CutoffUtc) { break }
        if ($item.EntryID) {
            [string] $item.EntryID
            $Limit--
        }
    }
}

function Get-TargetInbox {
    param($Context)

    $inbox = $Context.Store.GetDefaultFolder(6)
    if ($null -eq $inbox -or $inbox.StoreID -ne $Context.Store.StoreID) { throw 'TargetInboxUnavailable' }
    return $inbox
}

function Move-ArchiveBatch {
    param($Context, $Inbox, [string[]] $EntryIds, [datetime] $CutoffUtc, [long] $UsedBytes)

    if (-not (Get-AccountConnection -Context $Context).ArchiveAllowed) { throw 'ArchiveConnectionChanged' }
    $archivePath = $script:State.Pending.PstPath
    $destination = Get-PstFolder -Context $Context -Path $archivePath
    if (-not (Test-PrivateFile -Path $archivePath)) { throw 'PstStoreUnavailable' }
    if (-not $script:State.Pending.PstCreatedUtc) {
        $script:State.Pending.PstCreatedUtc = [IO.File]::GetCreationTimeUtc($archivePath).ToString('o')
        Save-State
    }
    $storeId = [string] $Context.Store.StoreID
    $inboxId = [string] $Inbox.EntryID
    $destinationStoreId = [string] $destination.StoreID
    $destinationId = [string] $destination.EntryID

    foreach ($entryId in $EntryIds) {
        try {
            $item = $Context.Session.GetItemFromID($entryId, $storeId)
        } catch {
            if ((Get-ErrorCode $_) -eq '0x8004010F') { continue } # Item already left the source store
            throw
        }
        $parent = $item.Parent
        if ($parent.EntryID -ne $inboxId -or $parent.StoreID -ne $storeId -or $item.Class -ne 43) { continue }
        $deliveryUtc = Get-MailDeliveryUtc -Item $item
        if ($deliveryUtc -ge $CutoffUtc) { continue }
        $itemSize = [long] $item.Size
        if ($itemSize -le 0) { throw 'InvalidMessageSize' }
        Assert-PstSpace -Path $archivePath -MessageBytes $itemSize
        $attachmentCount = $item.Attachments.Count
        # A stop between Move() and Save-State leaves an explicitly uncertain result.
        $script:State.Pending.InFlight = [pscustomobject] @{
            EntryId = $entryId
            StoreId = $storeId
            DeliveryUtc = $deliveryUtc.ToString('o')
            AttachmentCount = $attachmentCount
            StartedUtc = [DateTimeOffset]::UtcNow.ToString('o')
            ArchivedEntryId = $null
        }
        Save-State
        if (-not (Get-AccountConnection -Context $Context).ArchiveAllowed) {
            $script:State.Pending.InFlight = $null
            Save-State
            throw 'ArchiveConnectionChanged'
        }
        $currentParent = $item.Parent
        if ($currentParent.EntryID -ne $inboxId -or $currentParent.StoreID -ne $storeId) {
            $script:State.Pending.InFlight = $null
            Save-State
            continue
        }
        $moved = $item.Move($destination)
        if (-not $moved.EntryID) { throw 'PstDestinationVerificationFailed' }
        $script:State.Pending.InFlight.ArchivedEntryId = [string] $moved.EntryID
        Save-State
        $archived = $Context.Session.GetItemFromID([string] $moved.EntryID, $destinationStoreId)
        $movedParent = $archived.Parent
        if ($movedParent.StoreID -ne $destinationStoreId -or $movedParent.EntryID -ne $destinationId -or
            $archived.Attachments.Count -ne $attachmentCount -or (Get-MailDeliveryUtc -Item $archived) -ne $deliveryUtc) {
            throw 'PstDestinationVerificationFailed'
        }
        $script:State.Pending.MovedCount++
        $script:State.Pending.InFlight = $null
        $script:State.Pending.LastBatchBeforeBytes = $UsedBytes
        $script:State.Pending.LastMoveUtc = [DateTimeOffset]::UtcNow.ToString('o')
        Save-State
    }
}

function Resolve-UncertainMove {
    param($Context)

    $pending = $script:State.Pending
    if ($null -eq $pending -or -not $pending.InFlight -or $RunId -ne $pending.Id) { throw 'PendingMoveIdMismatch' }
    Assert-PendingPst -Pending $pending
    $move = $pending.InFlight
    if ($move.StoreId -ne $Context.Store.StoreID) { throw 'PendingMoveStoreMismatch' }
    $inbox = Get-TargetInbox -Context $Context
    $source = $null
    try {
        $source = $Context.Session.GetItemFromID($move.EntryId, $move.StoreId)
    } catch {
        if ((Get-ErrorCode $_) -ne '0x8004010F') { throw }
    }
    if ($MoveResolution -eq 'StillInInbox') {
        if ($null -eq $source -or $source.Class -ne 43 -or $source.Parent.EntryID -ne $inbox.EntryID -or
            $source.Parent.StoreID -ne $move.StoreId -or
            (Get-MailDeliveryUtc -Item $source).ToString('o') -ne $move.DeliveryUtc) {
            throw 'SourceMailCouldNotBeConfirmed'
        }
    } else {
        if ($null -ne $source) { throw 'SourceMailStillInExchangeStore' }
        $copyId = if ($ArchivedEntryId) { $ArchivedEntryId } else { $move.ArchivedEntryId }
        if (-not $copyId) { throw 'ArchivedEntryIdRequired' }
        if (-not (Test-PrivateFile -Path $pending.PstPath)) { throw 'PstStoreUnavailable' }
        $archive = Find-PstStore -Session $Context.Session -Path $pending.PstPath
        if ($null -eq $archive) { throw 'PstStoreUnavailable' }
        $copy = $Context.Session.GetItemFromID($copyId, $archive.StoreID)
        $root = $archive.GetRootFolder()
        if ($copy.Class -ne 43 -or $copy.Parent.StoreID -ne $archive.StoreID -or
            $copy.Parent.Name -ne $ArchiveFolderName -or $copy.Parent.Parent.EntryID -ne $root.EntryID -or
            $copy.Attachments.Count -ne $move.AttachmentCount -or
            (Get-MailDeliveryUtc -Item $copy).ToString('o') -ne $move.DeliveryUtc) {
            throw 'PstDestinationVerificationFailed'
        }
        $pending.MovedCount++
        $pending.LastMoveUtc = [DateTimeOffset]::UtcNow.ToString('o')
    }
    $pending.InFlight = $null
    $script:State.IssueReason = ''
    $script:State.IssueId = ''
    Save-State
    Write-Log -Name 'move_resolution_recorded' -Details @{ run_id = $pending.Id; resolution = $MoveResolution }
}

function Invoke-ArchiveCycle {
    param($Context, $Usage)

    $now = [DateTimeOffset]::UtcNow
    $pending = $script:State.Pending
    $cutoff = $now.UtcDateTime.AddHours(-$RetentionHours)
    if ($null -ne $pending) {
        Assert-PendingPst -Pending $pending
        $cutoff = [DateTimeOffset]::Parse($pending.CutoffUtc).UtcDateTime
        $details = [ordered] @{
            run_id = $pending.Id
            moved_count = $pending.MovedCount
            pst_path = $pending.PstPath
            before_bytes = $pending.BeforeBytes
            after_bytes = $Usage.UsedBytes
            cutoff_utc = $pending.CutoffUtc
            connection_mode = $Usage.Connection
            usage_source = $Usage.Source
        }
        if ($pending.InFlight) {
            Publish-Issue -Reason 'move_result_uncertain' -Details $details
            Write-Log -Name 'move_result_uncertain' -Details $details
            $script:ExitCode = 2
            return
        }
    }

    if ($null -eq $pending -and $Usage.UsedBytes -lt $TriggerBytes) {
        if ($script:State.IssueReason) {
            $script:State.IssueReason = ''
            $script:State.IssueId = ''
            Save-State
        }
        return
    }

    $inbox = Get-TargetInbox -Context $Context
    $entryIds = @(Get-ArchiveCandidates -Inbox $inbox -CutoffUtc $cutoff -Limit $MaxMessagesPerRun)
    if ($entryIds.Count -eq 0) {
        if ($null -eq $pending -or $pending.MovedCount -eq 0) {
            if ($Usage.UsedBytes -ge $TriggerBytes) {
                $details = @{ used_bytes = $Usage.UsedBytes; moved_count = 0 }
                Publish-Issue -Reason 'old_mail_insufficient' -Details $details
                Write-Log -Name 'old_mail_insufficient' -Details $details
            }
            if ($null -ne $pending) {
                $script:State.Pending = $null
                Save-State
            }
            return
        }

        $elapsed = $now - [DateTimeOffset]::Parse($pending.LastMoveUtc)
        if ($elapsed -lt $VerificationDelay -or
            ($Usage.UsedBytes -ge $TriggerBytes -and $elapsed -lt $VerificationTimeout)) {
            Write-Log -Name 'awaiting_usage_recheck' -Details $details
            return
        }
        if ($Usage.UsedBytes -lt $TriggerBytes) {
            if (-not (Get-AccountConnection -Context $Context).ArchiveAllowed) { throw 'ArchiveConnectionChanged' }
            Publish-Event -Id "archive-$($pending.Id)" -Type 'archive_completed' -Details $details
            Write-Log -Name 'archive_completed' -Details $details
            $script:State.IssueReason = ''
            $script:State.IssueId = ''
        } else {
            $details['usage_fell_since_last_batch'] = $Usage.UsedBytes -lt $pending.LastBatchBeforeBytes
            Publish-Issue -Reason 'old_mail_insufficient' -Details $details
            Write-Log -Name 'old_mail_insufficient' -Details $details
        }
        $script:State.Pending = $null
        Save-State
        return
    }

    if ($null -eq $pending) {
        $script:State.Pending = [pscustomobject] @{
            Id = [guid]::NewGuid().ToString('N')
            CutoffUtc = ([DateTimeOffset] $cutoff).ToString('o')
            BeforeBytes = $Usage.UsedBytes
            LastBatchBeforeBytes = $Usage.UsedBytes
            MovedCount = 0
            LastMoveUtc = $null
            InFlight = $null
            PstPath = Get-DatedPstPath -StartedUtc $now
            PstCreatedUtc = $null
        }
    }
    $beforeCount = $script:State.Pending.MovedCount
    Save-State
    Move-ArchiveBatch -Context $Context -Inbox $inbox -EntryIds $entryIds -CutoffUtc $cutoff -UsedBytes $Usage.UsedBytes
    Write-Log -Name 'batch_moved' -Details ([ordered] @{
        run_id = $script:State.Pending.Id
        moved_this_cycle = $script:State.Pending.MovedCount - $beforeCount
        moved_count = $script:State.Pending.MovedCount
        cutoff_utc = $cutoff.ToString('o')
        pst_path = $script:State.Pending.PstPath
    })
}

# ---------------------------------------------------------------------------
# Daily-once path. Stateless by design: no state.json, no Pending, no InFlight,
# no cross-run verification delay. Everything settles in this one invocation.
# The original Check/Preview/Archive/ResolveMove paths are left untouched.
# ---------------------------------------------------------------------------

function Invoke-DailyArchive {
    param($Context, $Usage)

    $now = [DateTimeOffset]::UtcNow
    $cutoff = $now.UtcDateTime.AddHours(-$RetentionHours)
    $inbox = Get-TargetInbox -Context $Context
    # No message-count cap: take every mail older than the retention cutoff.
    $entryIds = @(Get-ArchiveCandidates -Inbox $inbox -CutoffUtc $cutoff -Limit ([int]::MaxValue))
    $runId = [guid]::NewGuid().ToString('N')
    $pstPath = Get-DatedPstPath -StartedUtc $now
    $details = [ordered] @{
        run_id          = $runId
        eligible_count  = $entryIds.Count
        moved_count     = 0
        skipped_count   = 0
        pst_path        = $pstPath
        before_bytes    = $Usage.UsedBytes
        after_bytes     = $null
        cutoff_utc      = ([DateTimeOffset] $cutoff).ToString('o')
        connection_mode = $Usage.Connection
        usage_source    = $Usage.Source
    }
    $script:DailyDetails = $details

    if ($entryIds.Count -eq 0) {
        $d = [ordered] @{
            reason         = 'old_mail_insufficient'
            used_bytes     = $Usage.UsedBytes
            eligible_count = 0
            moved_count    = 0
            cutoff_utc     = $details['cutoff_utc']
        }
        Publish-Event -Id $runId -Type 'attention_required' -Details $d
        Write-Log -Name 'old_mail_insufficient' -Details $d
        return
    }

    $destination = Get-PstFolder -Context $Context -Path $pstPath
    if (-not (Test-PrivateFile -Path $pstPath)) { throw 'PstStoreUnavailable' }
    $storeId = [string] $Context.Store.StoreID
    $inboxId = [string] $inbox.EntryID
    $destinationStoreId = [string] $destination.StoreID
    $destinationId = [string] $destination.EntryID

    foreach ($entryId in $entryIds) {
        try {
            $item = $Context.Session.GetItemFromID($entryId, $storeId)
        } catch {
            if ((Get-ErrorCode $_) -eq '0x8004010F') {
                $details['skipped_count'] = [int] $details['skipped_count'] + 1
                continue
            }
            throw
        }
        $parent = $item.Parent
        if ($parent.EntryID -ne $inboxId -or $parent.StoreID -ne $storeId -or $item.Class -ne 43) {
            $details['skipped_count'] = [int] $details['skipped_count'] + 1
            continue
        }
        $deliveryUtc = Get-MailDeliveryUtc -Item $item
        if ($deliveryUtc -ge $cutoff) {
            $details['skipped_count'] = [int] $details['skipped_count'] + 1
            continue
        }
        $itemSize = [long] $item.Size
        if ($itemSize -le 0) { throw 'InvalidMessageSize' }
        Assert-PstSpace -Path $pstPath -MessageBytes $itemSize
        $attachmentCount = $item.Attachments.Count
        if (-not (Get-AccountConnection -Context $Context).ArchiveAllowed) { throw 'ArchiveConnectionChanged' }
        $currentParent = $item.Parent
        if ($currentParent.EntryID -ne $inboxId -or $currentParent.StoreID -ne $storeId) {
            $details['skipped_count'] = [int] $details['skipped_count'] + 1
            continue
        }
        $moved = $item.Move($destination)
        if (-not $moved.EntryID) { throw 'PstDestinationVerificationFailed' }
        $archived = $Context.Session.GetItemFromID([string] $moved.EntryID, $destinationStoreId)
        $movedParent = $archived.Parent
        if ($movedParent.StoreID -ne $destinationStoreId -or $movedParent.EntryID -ne $destinationId -or
            $archived.Attachments.Count -ne $attachmentCount -or
            (Get-MailDeliveryUtc -Item $archived) -ne $deliveryUtc) {
            throw 'PstDestinationVerificationFailed'
        }
        $details['moved_count'] = [int] $details['moved_count'] + 1
    }

    $after = Get-Usage -Context $Context
    $details['after_bytes'] = $after.UsedBytes
    Publish-Event -Id "archive-$runId" -Type 'archive_completed' -Details $details
    Write-Log -Name 'archive_completed' -Details $details

    # Cached-mode usage can lag behind the server. Report it as an attention
    # item instead of silently claiming the mailbox is back under the line.
    if ($after.UsedBytes -ge $TriggerBytes) {
        $d = [ordered] @{
            reason      = 'usage_still_above_trigger_after_archive'
            used_bytes  = $after.UsedBytes
            moved_count = $details['moved_count']
            pst_path    = $pstPath
        }
        Publish-Event -Id $runId -Type 'attention_required' -Details $d
        Write-Log -Name 'usage_still_above_trigger_after_archive' -Details $d
    }
}

function Invoke-DailyOnceRun {
    Initialize-DataDirectory
    $script:DailyDetails = $null
    $lockPath = Join-Path $DataDirectory 'monitor.lock'
    [void] (Test-PrivateFile -Path $lockPath)
    try {
        $lock = [IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None')
    } catch [IO.IOException] {
        $code = $_.Exception.GetBaseException().HResult -band 0xffff
        if ($code -ne 32) { throw }
        Write-Output '{"action":"overlapping_run_skipped"}'
        $script:ExitCode = 3
        return
    }

    try {
        $context = Get-MailboxContext
        $usage = Get-Usage -Context $context
        Write-Log -Name 'usage_sample' -Details ([ordered] @{
            used_bytes          = $usage.UsedBytes
            quota_bytes         = $QuotaBytes
            trigger_bytes       = $TriggerBytes
            send_quota_bytes    = $usage.SendQuotaBytes
            connection_mode     = $usage.Connection
            connected           = $usage.Connected
            archive_allowed     = $usage.ArchiveAllowed
            at_or_above_trigger = $usage.UsedBytes -ge $TriggerBytes
            usage_source        = $usage.Source
        })
        if (-not $usage.Connected) { throw 'OutlookExchangeDisconnected' }
        if ($usage.UsedBytes -lt $TriggerBytes) {
            Write-Output '{"action":"below_trigger"}'
            return
        }
        if (-not $usage.ArchiveAllowed) { throw 'ArchiveConnectionChanged' }
        Invoke-DailyArchive -Context $context -Usage $usage
    } catch {
        $details = [ordered] @{
            reason     = 'archive_operation_failed'
            error_code = (Get-ErrorCode $_)
            cause      = [string] $_.Exception.Message
        }
        if ($null -ne $script:DailyDetails) {
            $details['moved_count']     = $script:DailyDetails['moved_count']
            $details['eligible_count']  = $script:DailyDetails['eligible_count']
            $details['pst_path']        = $script:DailyDetails['pst_path']
            $details['partial_archive'] = $true
        }
        $id = [guid]::NewGuid().ToString('N')
        try { Publish-Event -Id $id -Type 'attention_required' -Details $details } catch { }
        try { Write-Log -Name 'archive_operation_failed' -Details $details } catch { }
        $out = [ordered] @{ action = 'archive_operation_failed'; details = $details }
        Write-Output ($out | ConvertTo-Json -Compress -Depth 6)
        $script:ExitCode = 1
    } finally {
        $lock.Dispose()
    }
}

function Invoke-Monitor {
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or
        [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        throw 'Run with Windows PowerShell 5.1 (powershell.exe -STA).'
    }
    if ($Mode -eq 'Archive' -and $DailyOnce) {
        Invoke-DailyOnceRun
        return
    }
    if ($Mode -eq 'ResolveMove' -and (-not $RunId -or -not $MoveResolution)) {
        throw 'ResolveMove requires -RunId and -MoveResolution.'
    }
    Initialize-DataDirectory
    $lockPath = Join-Path $DataDirectory 'monitor.lock'
    [void] (Test-PrivateFile -Path $lockPath)
    try {
        $lock = [IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None')
    } catch [IO.IOException] {
        $code = $_.Exception.GetBaseException().HResult -band 0xffff
        if ($code -ne 32) { throw }
        Write-Output '{"action":"overlapping_run_skipped"}'
        return
    }

    try {
        if ($Mode -in @('Archive', 'ResolveMove')) {
            try {
                $script:State = Read-State
            } catch {
                $details = @{ reason = 'state_invalid_or_unreadable'; error_code = (Get-ErrorCode $_) }
                Write-Log -Name 'state_invalid_or_unreadable' -Details $details
                $script:ExitCode = 1
                return
            }
        }

        try {
            $context = Get-MailboxContext
            $usage = Get-Usage -Context $context
            Write-Log -Name 'usage_sample' -Details ([ordered] @{
                used_bytes = $usage.UsedBytes
                quota_bytes = $QuotaBytes
                trigger_bytes = $TriggerBytes
                send_quota_bytes = $usage.SendQuotaBytes
                connection_mode = $usage.Connection
                connected = $usage.Connected
                archive_allowed = $usage.ArchiveAllowed
                at_or_above_trigger = $usage.UsedBytes -ge $TriggerBytes
                usage_source = $usage.Source
            })
            if (-not $usage.Connected) { throw 'OutlookExchangeDisconnected' }
        } catch {
            $knownReasons = @('AmbiguousMailbox', 'MailboxNotConfigured', 'MailboxIsNotExchange',
                'ExchangeDeliveryStoreUnavailable', 'InvalidMailboxUsage', 'OutlookExchangeDisconnected')
            $reason = $_.Exception.Message
            if ($reason -notin $knownReasons) { $reason = 'outlook_or_usage_unavailable' }
            $details = @{ reason = $reason; error_code = (Get-ErrorCode $_) }
            Write-Log -Name 'monitor_unavailable' -Details $details
            if ($Mode -eq 'Archive') {
                $script:State.UnavailableCount++
                Save-State
                if ($script:State.UnavailableCount -ge 3) {
                    Publish-Issue -Reason 'monitor_unavailable' -Details @{
                        cause = $reason
                        consecutive_checks = $script:State.UnavailableCount
                    }
                }
            }
            $script:ExitCode = 2
            return
        }

        if ($Mode -eq 'Preview') {
            $cutoff = [datetime]::UtcNow.AddHours(-$RetentionHours)
            $inbox = Get-TargetInbox -Context $context
            $candidates = @(Get-ArchiveCandidates -Inbox $inbox -CutoffUtc $cutoff -Limit $MaxMessagesPerRun)
            Write-Log -Name 'archive_preview' -Details ([ordered] @{
                cutoff_utc = $cutoff.ToString('o')
                eligible_up_to_limit = $candidates.Count
                batch_limit = $MaxMessagesPerRun
            })
        }

        if ($Mode -eq 'ResolveMove') {
            if (-not $usage.ArchiveAllowed) { throw 'ResolveMove needs an eligible Exchange connection.' }
            Resolve-UncertainMove -Context $context
        }

        if ($Mode -eq 'Archive') {
            if ($script:State.UnavailableCount -gt 0) {
                $script:State.UnavailableCount = 0
                if ($script:State.IssueReason -eq 'monitor_unavailable') {
                    $script:State.IssueReason = ''
                    $script:State.IssueId = ''
                }
                Save-State
                Write-Log -Name 'monitor_recovered'
            }
            try {
                if (-not $usage.ArchiveAllowed -and ($usage.UsedBytes -ge $TriggerBytes -or $null -ne $script:State.Pending)) {
                    $details = @{ connection_mode = $usage.Connection; used_bytes = $usage.UsedBytes; usage_source = $usage.Source }
                    Publish-Issue -Reason 'archive_connection_not_eligible' -Details $details
                    Write-Log -Name 'archive_connection_not_eligible' -Details $details
                    $script:ExitCode = 2
                    return
                }
                Invoke-ArchiveCycle -Context $context -Usage $usage
            } catch {
                $details = @{ error_code = (Get-ErrorCode $_); used_bytes = $usage.UsedBytes }
                $knownCause = $_.Exception.Message
                if ($knownCause -in @('PstDiskSpaceLow', 'PstCapacityNearLimit', 'PendingPstMissingOrReplaced',
                        'ArchiveConnectionChanged', 'PstDestinationVerificationFailed', 'InvalidDeliveryTime', 'InvalidMessageSize')) {
                    $details['cause'] = $knownCause
                }
                if ($null -ne $script:State.Pending) {
                    $details['moved_count'] = $script:State.Pending.MovedCount
                    $details['run_id'] = $script:State.Pending.Id
                    $details['pst_path'] = $script:State.Pending.PstPath
                }
                $reason = if ($null -ne $script:State.Pending -and $script:State.Pending.InFlight) {
                    'move_result_uncertain'
                } else { 'archive_operation_failed' }
                Publish-Issue -Reason $reason -Details $details
                Write-Log -Name $reason -Details $details
                $script:ExitCode = 1
            }
        }
    } finally {
        $lock.Dispose()
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    try {
        Invoke-Monitor
    } catch {
        Write-Error -ErrorAction Continue -Message ('MailboxMonitor stopped. ' + $_.Exception.Message)
        $script:ExitCode = 1
    }
    exit $script:ExitCode
}
