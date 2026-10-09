#Requires -Version 5.1
<#
    HFM Extended Analytics Template Manager (single-file build 2026-10-09 09:54).

    Manage HFM Extended Analytics templates (the SS_PREF* rows of <APP>_USERPARAMS):
    browse, edit, create, rename, delete, copy between users/applications, export/import, back up and restore.

    Start:  right-click > Run with PowerShell
       or:  powershell -ExecutionPolicy Bypass -File EATemplateManager.ps1

    Needs Windows PowerShell 5.1 and .NET Framework 4.x (standard on Windows Server). Nothing to install.
    Your connection details and last selection (never the password) are remembered per Windows user in
    %APPDATA%\EATemplateManager\preferences.json. Backups\ and EATemplateManager.log are created next to this file.

    Built from separate source files by Build-SingleFile.ps1; edit those, not this file.
#>

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`""
    return
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
#region Embedded data module (source: EATemplate.psm1)
$null = New-Module -Name EATemplate -ScriptBlock {
<#
    EATemplate.psm1 - data layer for HFM Extended Analytics (EA) templates.

    EA templates live in <APP>_USERPARAMS:
      Username nvarchar(128), EntryIDX int, ParameterKey nvarchar(60), BLOBData varbinary(2000)
    - Payloads are UTF-16LE text (no BOM, no terminator), split into 2000-byte chunks (EntryIDX 1..n).
    - SS_PREF            : per-user index  <SS_TEMPLATES><SS_TEMPLATE name="X" key="SS_PREF_X"/>...</SS_TEMPLATES>
    - SS_PREF_<name>     : one template    <povTemplate version="11.2"><povEA><![CDATA[...]]></povEA><options>...</options></povTemplate>
    - SS_DEFAULT         : last-used EA settings (same format as a template)
#>

Set-StrictMode -Version 2

$script:Conn         = $null
$script:KnownApps    = @()
$script:ChunkSize    = 2000
$script:IndexKey     = 'SS_PREF'
$script:DefaultKey   = 'SS_DEFAULT'
$script:KeyPrefix    = 'SS_PREF_'
$script:MaxKeyLength = 60
$script:IndexHeader  = "<?xml version = '1.0' encoding = 'UTF-16LE'?>`r`n"
$script:BackupDir    = Join-Path $PSScriptRoot 'Backups'
$script:DimPattern   = '(S|Y|P|W|E|V|A|I|C\d+)[#{]'
$script:RequiredDims = @('S', 'Y', 'P', 'W', 'E', 'V', 'A', 'I')
$script:DimNames     = @{ S = 'Scenario'; Y = 'Year'; P = 'Period'; W = 'View'; E = 'Entity'; V = 'Value'; A = 'Account'; I = 'ICP' }

$script:NewTemplateXml = '<povTemplate version="11.2"><povEA><![CDATA[]]></povEA><options>' +
    '<tablePrefix></tablePrefix><filePrefix></filePrefix><exportOption>0</exportOption><extractFormat>0</extractFormat>' +
    '<selectedDSN></selectedDSN><extractDynAccts>false</extractDynAccts><extractData>true</extractData>' +
    '<extractCalculatedData>true</extractCalculatedData><extractDerivedData>true</extractDerivedData>' +
    '<extractPhasedSubmission>false</extractPhasedSubmission><extractCellText>false</extractCellText>' +
    '<lineItemDetailOption>EA_LINEITEM_EXCLUDE</lineItemDetailOption><delimiter>;</delimiter></options></povTemplate>'

#region Connection and SQL helpers

function Connect-EaDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Database,
        [string]$UserName,
        [securestring]$Password
    )
    Disconnect-EaDatabase
    $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $csb['Data Source'] = $Server
    $csb['Initial Catalog'] = $Database
    $csb['Application Name'] = 'EATemplateManager'
    $conn = New-Object System.Data.SqlClient.SqlConnection
    if ($UserName) {
        $conn.ConnectionString = $csb.ConnectionString
        $pw = $Password.Copy()
        $pw.MakeReadOnly()
        $conn.Credential = New-Object System.Data.SqlClient.SqlCredential($UserName, $pw)
    }
    else {
        $csb['Integrated Security'] = $true
        $conn.ConnectionString = $csb.ConnectionString
    }
    $conn.Open()
    $script:Conn = $conn
    $script:KnownApps = @()
    [void](Get-HfmApp)
}

function Disconnect-EaDatabase {
    if ($script:Conn) {
        try { $script:Conn.Dispose() } catch { }
        $script:Conn = $null
    }
}

function Test-EaConnection {
    return [bool]($script:Conn -and $script:Conn.State -eq [System.Data.ConnectionState]::Open)
}

function Assert-EaConnection {
    if (-not $script:Conn) { throw 'Not connected to the HFM database.' }
    if ($script:Conn.State -ne [System.Data.ConnectionState]::Open) {
        try { $script:Conn.Close() } catch { }
        $script:Conn.Open()
    }
}

function New-EaCommand {
    param([string]$Sql, [hashtable]$Parameters, $Transaction)
    Assert-EaConnection
    $cmd = $script:Conn.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = 60
    if ($Transaction) { $cmd.Transaction = $Transaction }
    if ($Parameters) {
        foreach ($k in $Parameters.Keys) {
            $v = $Parameters[$k]
            if ($null -eq $v) {
                [void]$cmd.Parameters.AddWithValue("@$k", [DBNull]::Value)
            }
            elseif ($v -is [byte[]]) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::VarBinary, $script:ChunkSize)
                $p.Value = $v
            }
            elseif ($v -is [string]) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::NVarChar, 4000)
                $p.Value = $v
            }
            else {
                [void]$cmd.Parameters.AddWithValue("@$k", $v)
            }
        }
    }
    return $cmd
}

function Invoke-EaQuery {
    param([string]$Sql, [hashtable]$Parameters, $Transaction)
    $cmd = New-EaCommand -Sql $Sql -Parameters $Parameters -Transaction $Transaction
    try {
        $dt = New-Object System.Data.DataTable
        $reader = $cmd.ExecuteReader()
        try { $dt.Load($reader) } finally { $reader.Dispose() }
        return , $dt
    }
    finally { $cmd.Dispose() }
}

function Invoke-EaNonQuery {
    param([string]$Sql, [hashtable]$Parameters, $Transaction)
    $cmd = New-EaCommand -Sql $Sql -Parameters $Parameters -Transaction $Transaction
    try { return $cmd.ExecuteNonQuery() } finally { $cmd.Dispose() }
}

function Invoke-EaTransaction {
    # Runs $Body with a SqlTransaction argument; commits on success, rolls back on any error.
    param([Parameter(Mandatory)][scriptblock]$Body)
    Assert-EaConnection
    $tx = $script:Conn.BeginTransaction()
    try {
        $result = & $Body $tx
        $tx.Commit()
        return $result
    }
    catch {
        try { $tx.Rollback() } catch { }
        throw
    }
    finally { $tx.Dispose() }
}

function Get-ParamTable {
    param([Parameter(Mandatory)][string]$App)
    if ($App -notmatch '^[A-Za-z0-9_]+$') { throw "Invalid application name '$App'." }
    if ($script:KnownApps -notcontains $App) { [void](Get-HfmApp) }
    if ($script:KnownApps -notcontains $App) { throw "Unknown HFM application '$App'." }
    return "[${App}_USERPARAMS]"
}

#endregion

#region Discovery

function Get-HfmApp {
    # HFM applications that have a USERPARAMS table.
    $apps = @()
    try {
        $dt = Invoke-EaQuery -Sql "SELECT APPNAME FROM HSX_DATASOURCES WHERE PRODUCT = 'Financial Management' AND OBJECT_ID(APPNAME + '_USERPARAMS', 'U') IS NOT NULL ORDER BY APPNAME"
        $apps = @($dt.Rows | ForEach-Object { [string]$_.APPNAME })
    }
    catch { }
    if ($apps.Count -eq 0) {
        $dt = Invoke-EaQuery -Sql "SELECT LEFT(name, LEN(name) - 11) AS APPNAME FROM sys.tables WHERE name LIKE '%[_]USERPARAMS' ORDER BY name"
        $apps = @($dt.Rows | ForEach-Object { [string]$_.APPNAME })
    }
    $script:KnownApps = @($apps | Where-Object { $_ -match '^[A-Za-z0-9_]+$' })
    return $script:KnownApps
}

function Get-EaDsn {
    # Only the DSN names are read; HSX_DSN_DETAILS also holds (encrypted) credentials.
    try {
        $dt = Invoke-EaQuery -Sql 'SELECT DsnName FROM HSX_DSN_DETAILS ORDER BY DsnName'
        return @($dt.Rows | ForEach-Object { [string]$_.DsnName })
    }
    catch { return @() }
}

function Get-EaTemplateUser {
    # Users that have EA templates (or EA defaults) in the application.
    param([Parameter(Mandatory)][string]$App)
    $t = Get-ParamTable $App
    $dt = Invoke-EaQuery -Sql "SELECT DISTINCT Username FROM $t WHERE ParameterKey LIKE 'SS[_]PREF%' OR ParameterKey = 'SS_DEFAULT' ORDER BY Username"
    return @($dt.Rows | ForEach-Object { [string]$_.Username })
}

function Get-HfmUser {
    # All users known to HFM (activity table) plus anyone with rows in the application's USERPARAMS.
    param([string]$App)
    $users = New-Object 'System.Collections.Generic.SortedSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    try {
        $dt = Invoke-EaQuery -Sql 'SELECT SUSERNAME FROM HSV_ACTIVITY_USERS'
        foreach ($r in $dt.Rows) { [void]$users.Add([string]$r.SUSERNAME) }
    }
    catch { }
    if ($App) {
        $t = Get-ParamTable $App
        $dt = Invoke-EaQuery -Sql "SELECT DISTINCT Username FROM $t"
        foreach ($r in $dt.Rows) { [void]$users.Add([string]$r.Username) }
    }
    return [string[]]@(foreach ($u in $users) { $u })
}

function Get-HfmActiveUser {
    # Users with an active HFM session on the application (their cached preferences may overwrite our changes).
    param([Parameter(Mandatory)][string]$App)
    try {
        $dt = Invoke-EaQuery -Sql 'SELECT DISTINCT u.SUSERNAME FROM HSV_ACTIVITY_SESSIONS s JOIN HSV_ACTIVITY_USERS u ON u.LUSERID = s.LUSERID WHERE s.SAPPNAME = @app' -Parameters @{ app = $App }
        return @($dt.Rows | ForEach-Object { [string]$_.SUSERNAME })
    }
    catch { return @() }
}

#endregion

#region Raw USERPARAMS access

function Read-UserParamBytes {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key, $Transaction)
    $t = Get-ParamTable $App
    $dt = Invoke-EaQuery -Sql "SELECT BLOBData FROM $t WHERE Username = @u AND ParameterKey = @k ORDER BY EntryIDX" -Parameters @{ u = $User; k = $Key } -Transaction $Transaction
    if ($dt.Rows.Count -eq 0) { return $null }
    $ms = New-Object System.IO.MemoryStream
    foreach ($r in $dt.Rows) {
        $b = [byte[]]$r.BLOBData
        $ms.Write($b, 0, $b.Length)
    }
    return , $ms.ToArray()
}

function Read-UserParam {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key, $Transaction)
    $bytes = Read-UserParamBytes -App $App -User $User -Key $Key -Transaction $Transaction
    if ($null -eq $bytes) { return $null }
    return [System.Text.Encoding]::Unicode.GetString($bytes)
}

function Write-UserParamBytes {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][byte[]]$Bytes, [Parameter(Mandatory)]$Transaction)
    if ($Key.Length -gt $script:MaxKeyLength) { throw "Parameter key '$Key' exceeds $($script:MaxKeyLength) characters." }
    $t = Get-ParamTable $App
    [void](Invoke-EaNonQuery -Sql "DELETE FROM $t WHERE Username = @u AND ParameterKey = @k" -Parameters @{ u = $User; k = $Key } -Transaction $Transaction)
    $idx = 1
    for ($offset = 0; $offset -lt $Bytes.Length; $offset += $script:ChunkSize) {
        $len = [Math]::Min($script:ChunkSize, $Bytes.Length - $offset)
        $chunk = New-Object byte[] $len
        [Array]::Copy($Bytes, $offset, $chunk, 0, $len)
        [void](Invoke-EaNonQuery -Sql "INSERT INTO $t (Username, EntryIDX, ParameterKey, BLOBData) VALUES (@u, @i, @k, @b)" -Parameters @{ u = $User; i = $idx; k = $Key; b = $chunk } -Transaction $Transaction)
        $idx++
    }
}

function Write-UserParam {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Text, [Parameter(Mandatory)]$Transaction)
    Write-UserParamBytes -App $App -User $User -Key $Key -Bytes ([System.Text.Encoding]::Unicode.GetBytes($Text)) -Transaction $Transaction
}

function Remove-UserParam {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)]$Transaction)
    $t = Get-ParamTable $App
    [void](Invoke-EaNonQuery -Sql "DELETE FROM $t WHERE Username = @u AND ParameterKey = @k" -Parameters @{ u = $User; k = $Key } -Transaction $Transaction)
}

function Get-UserParamKey {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, $Transaction)
    $t = Get-ParamTable $App
    $dt = Invoke-EaQuery -Sql "SELECT DISTINCT ParameterKey FROM $t WHERE Username = @u AND ParameterKey LIKE 'SS[_]PREF[_]%'" -Parameters @{ u = $User } -Transaction $Transaction
    return @($dt.Rows | ForEach-Object { [string]$_.ParameterKey })
}

#endregion

#region Backups

function Backup-EaRows {
    <#
        Dumps USERPARAMS rows to Backups\<timestamp>_<app>_<tag>.csv (hex BLOBs).
        -Targets: objects with User and Key; omit to back up every EA row (SS_PREF*, SS_DEFAULT) in the app.
        Keys that had no rows are recorded with EntryIDX 0 so a restore can remove them again.
    #>
    param([Parameter(Mandatory)][string]$App, [object[]]$Targets, [string]$Tag = 'backup')
    $t = Get-ParamTable $App
    $rows = New-Object System.Collections.Generic.List[object]
    if ($Targets) {
        $seen = @{}
        foreach ($tg in $Targets) {
            $id = "$($tg.User)|$($tg.Key)".ToUpperInvariant()
            if ($seen.ContainsKey($id)) { continue }
            $seen[$id] = $true
            $dt = Invoke-EaQuery -Sql "SELECT Username, EntryIDX, ParameterKey, BLOBData FROM $t WHERE Username = @u AND ParameterKey = @k ORDER BY EntryIDX" -Parameters @{ u = $tg.User; k = $tg.Key }
            if ($dt.Rows.Count -eq 0) {
                $rows.Add([pscustomobject]@{ App = $App; Username = $tg.User; EntryIDX = 0; ParameterKey = $tg.Key; BLOBHex = '' })
            }
            foreach ($r in $dt.Rows) { $rows.Add((ConvertTo-BackupRow $App $r)) }
        }
    }
    else {
        $dt = Invoke-EaQuery -Sql "SELECT Username, EntryIDX, ParameterKey, BLOBData FROM $t WHERE ParameterKey LIKE 'SS[_]PREF%' OR ParameterKey = 'SS_DEFAULT' ORDER BY Username, ParameterKey, EntryIDX"
        foreach ($r in $dt.Rows) { $rows.Add((ConvertTo-BackupRow $App $r)) }
    }
    if (-not (Test-Path $script:BackupDir)) { [void](New-Item -ItemType Directory -Path $script:BackupDir) }
    $safeTag = $Tag -replace '[^A-Za-z0-9_-]', '_'
    $path = Join-Path $script:BackupDir ('{0}_{1}_{2}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $App, $safeTag)
    $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    return $path
}

function ConvertTo-BackupRow {
    param([string]$App, $Row)
    $b = [byte[]]$Row.BLOBData
    return [pscustomobject]@{
        App          = $App
        Username     = [string]$Row.Username
        EntryIDX     = [int]$Row.EntryIDX
        ParameterKey = [string]$Row.ParameterKey
        BLOBHex      = [BitConverter]::ToString($b).Replace('-', '')
    }
}

function Restore-EaBackup {
    <# Restores every (App, User, Key) in a backup CSV to the backed-up state, in one transaction. #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([Parameter(Mandatory)][string]$Path)
    $rows = @(Import-Csv -Path $Path)
    if ($rows.Count -eq 0) { throw "Backup '$Path' is empty." }
    $groups = $rows | Group-Object App, Username, ParameterKey
    if (-not $PSCmdlet.ShouldProcess("$($groups.Count) parameter key(s) from $Path", 'Restore')) { return }
    $apps = @($rows | Select-Object -ExpandProperty App -Unique)
    foreach ($a in $apps) {
        $pre = @($rows | Where-Object { $_.App -eq $a } | ForEach-Object { [pscustomobject]@{ User = $_.Username; Key = $_.ParameterKey } })
        [void](Backup-EaRows -App $a -Targets $pre -Tag 'pre-restore')
    }
    Invoke-EaTransaction {
        param($tx)
        foreach ($g in $groups) {
            $first = $g.Group[0]
            $t = Get-ParamTable $first.App
            [void](Invoke-EaNonQuery -Sql "DELETE FROM $t WHERE Username = @u AND ParameterKey = @k" -Parameters @{ u = $first.Username; k = $first.ParameterKey } -Transaction $tx)
            foreach ($r in ($g.Group | Where-Object { [int]$_.EntryIDX -gt 0 } | Sort-Object { [int]$_.EntryIDX })) {
                $hex = $r.BLOBHex
                $bytes = New-Object byte[] ($hex.Length / 2)
                for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [Convert]::ToByte($hex.Substring($i * 2, 2), 16) }
                [void](Invoke-EaNonQuery -Sql "INSERT INTO $t (Username, EntryIDX, ParameterKey, BLOBData) VALUES (@u, @i, @k, @b)" -Parameters @{ u = $r.Username; i = [int]$r.EntryIDX; k = $r.ParameterKey; b = $bytes } -Transaction $tx)
            }
        }
    }
}

#endregion

#region Template index (SS_PREF)

function ConvertFrom-EaIndexXml {
    param([string]$Text)
    $list = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Text)) { return , $list }
    $body = $Text -replace '^\s*<\?xml[^>]*\?>\s*', ''
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($body)
    foreach ($n in $doc.SelectNodes('/SS_TEMPLATES/SS_TEMPLATE')) {
        $list.Add([pscustomobject]@{ Name = $n.GetAttribute('name'); Key = $n.GetAttribute('key') })
    }
    return , $list
}

function ConvertTo-EaIndexXml {
    param([object[]]$Entries)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($script:IndexHeader).Append('<SS_TEMPLATES>')
    foreach ($e in $Entries) {
        [void]$sb.Append('<SS_TEMPLATE name="').Append((ConvertTo-XmlAttr $e.Name)).Append('" key="').Append((ConvertTo-XmlAttr $e.Key)).Append('"/>')
    }
    [void]$sb.Append('</SS_TEMPLATES>')
    return $sb.ToString()
}

function ConvertTo-XmlAttr {
    param([string]$Value)
    return $Value.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Get-EaIndex {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, $Transaction)
    $text = Read-UserParam -App $App -User $User -Key $script:IndexKey -Transaction $Transaction
    return , (ConvertFrom-EaIndexXml $text)
}

function Set-EaIndex {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [AllowEmptyCollection()][object[]]$Entries, [Parameter(Mandatory)]$Transaction)
    Write-UserParam -App $App -User $User -Key $script:IndexKey -Text (ConvertTo-EaIndexXml $Entries) -Transaction $Transaction
}

function Get-EaTemplateList {
    <#
        Merges the SS_PREF index with the SS_PREF_* rows actually present.
        Status: OK | NotInIndex (row exists, index entry missing) | MissingData (index entry, no row)
    #>
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User)
    $index = Get-EaIndex -App $App -User $User
    $keys = Get-UserParamKey -App $App -User $User
    $keySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($k in $keys) { [void]$keySet.Add($k) }
    $listed = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($e in $index) {
        [void]$listed.Add($e.Key)
        $status = if ($keySet.Contains($e.Key)) { 'OK' } else { 'MissingData' }
        $result.Add([pscustomobject]@{ Name = $e.Name; Key = $e.Key; Status = $status })
    }
    foreach ($k in ($keys | Sort-Object)) {
        if (-not $listed.Contains($k)) {
            $result.Add([pscustomobject]@{ Name = $k.Substring($script:KeyPrefix.Length); Key = $k; Status = 'NotInIndex' })
        }
    }
    return , $result
}

function Repair-EaIndex {
    <# Adds orphan SS_PREF_* rows to the index and drops index entries without data. #>
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User)
    [void](Backup-EaRows -App $App -Targets @([pscustomobject]@{ User = $User; Key = $script:IndexKey }) -Tag 'repair-index')
    Invoke-EaTransaction {
        param($tx)
        $index = Get-EaIndex -App $App -User $User -Transaction $tx
        $keys = Get-UserParamKey -App $App -User $User -Transaction $tx
        $keySet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($k in $keys) { [void]$keySet.Add($k) }
        $entries = New-Object System.Collections.Generic.List[object]
        $listed = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $index) {
            if ($keySet.Contains($e.Key) -and $listed.Add($e.Key)) { $entries.Add($e) }
        }
        foreach ($k in ($keys | Sort-Object)) {
            if ($listed.Add($k)) { $entries.Add([pscustomobject]@{ Name = $k.Substring($script:KeyPrefix.Length); Key = $k }) }
        }
        Set-EaIndex -App $App -User $User -Entries $entries.ToArray() -Transaction $tx
    }
}

#endregion

#region Template XML model

function Get-EaTemplateKey {
    param([Parameter(Mandatory)][string]$Name)
    return $script:KeyPrefix + $Name
}

function Test-EaTemplateName {
    # Returns an error message, or $null when the name is valid.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Template name is required.' }
    if ($Name -ne $Name.Trim()) { return 'Template name must not start or end with spaces.' }
    if ($Name -match '["<>&]') { return 'Template name must not contain " < > or &.' }
    $max = $script:MaxKeyLength - $script:KeyPrefix.Length
    if ($Name.Length -gt $max) { return "Template name must be $max characters or fewer." }
    return $null
}

function Get-NewEaTemplateXml {
    return $script:NewTemplateXml
}

function ConvertFrom-EaTemplateXml {
    <# Returns @{ Version; Pov; Options (ordered name->value) } for a povTemplate XML string. #>
    param([Parameter(Mandatory)][string]$Xml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($Xml)
    $root = $doc.DocumentElement
    if ($root.Name -ne 'povTemplate') { throw "Not an EA template: root element is <$($root.Name)>, expected <povTemplate>." }
    $povNode = $root.SelectSingleNode('povEA')
    $options = [ordered]@{}
    $optNode = $root.SelectSingleNode('options')
    if ($optNode) {
        foreach ($c in $optNode.ChildNodes) {
            if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element) { $options[$c.Name] = $c.InnerText }
        }
    }
    return [pscustomobject]@{
        Version = $root.GetAttribute('version')
        Pov     = if ($povNode) { $povNode.InnerText } else { '' }
        Options = $options
    }
}

function ConvertTo-EaTemplateXml {
    <#
        Applies a POV string and option values onto a base template XML and returns compact XML.
        Unknown elements and element order in the base are preserved; missing options are appended.
    #>
    param([Parameter(Mandatory)][string]$BaseXml, [Parameter(Mandatory)][AllowEmptyString()][string]$Pov, [System.Collections.IDictionary]$Options)
    if ($Pov.Contains(']]>')) { throw 'POV must not contain "]]>".' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($BaseXml)
    $root = $doc.DocumentElement
    $povNode = $root.SelectSingleNode('povEA')
    if (-not $povNode) {
        $povNode = $doc.CreateElement('povEA')
        [void]$root.PrependChild($povNode)
    }
    $cdata = $povNode.FirstChild
    if ($povNode.ChildNodes.Count -eq 1 -and $cdata.NodeType -eq [System.Xml.XmlNodeType]::CDATA) {
        $cdata.Value = $Pov
    }
    else {
        $povNode.RemoveAll()
        [void]$povNode.AppendChild($doc.CreateCDataSection($Pov))
    }
    if ($Options) {
        $optNode = $root.SelectSingleNode('options')
        if (-not $optNode) {
            $optNode = $doc.CreateElement('options')
            [void]$root.AppendChild($optNode)
        }
        foreach ($name in $Options.Keys) {
            $el = $optNode.SelectSingleNode($name)
            if (-not $el) {
                $el = $doc.CreateElement($name)
                [void]$optNode.AppendChild($el)
            }
            $value = [string]$Options[$name]
            if ($el.InnerText -ne $value) {
                $el.InnerText = $value
                if ($el -is [System.Xml.XmlElement]) { $el.IsEmpty = $false }
            }
        }
    }
    return $doc.OuterXml
}

#endregion

#region POV parsing

function Split-EaPov {
    <#
        Splits a POV such as  S#Actual.Y#2015.P#Jan;Feb.E{Top.[Hierarchy]}.C1{[Base]}
        into ordered [pscustomobject]@{ Dim; Expr } items (Expr keeps its leading # or {).
        A '.' separates dimensions only when outside {} [] <> and followed by a dimension prefix,
        so Parent.Child members and member-list functions are kept intact.
    #>
    param([AllowEmptyString()][string]$Pov)
    $items = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($Pov)) { return , $items }
    $segments = New-Object System.Collections.Generic.List[string]
    $depth = 0
    $start = 0
    $boundary = New-Object System.Text.RegularExpressions.Regex ('\G' + $script:DimPattern)
    for ($i = 0; $i -lt $Pov.Length; $i++) {
        $ch = $Pov[$i]
        if ('{[<'.IndexOf($ch) -ge 0) { $depth++ }
        elseif ('}]>'.IndexOf($ch) -ge 0) { if ($depth -gt 0) { $depth-- } }
        elseif ($ch -eq '.' -and $depth -eq 0 -and $boundary.IsMatch($Pov, $i + 1)) {
            $segments.Add($Pov.Substring($start, $i - $start))
            $start = $i + 1
        }
    }
    $segments.Add($Pov.Substring($start))
    foreach ($seg in $segments) {
        $m = [regex]::Match($seg, '^(S|Y|P|W|E|V|A|I|C\d+)(?=[#{])')
        if (-not $m.Success) { throw "Cannot parse POV segment '$seg'." }
        $items.Add([pscustomobject]@{ Dim = $m.Groups[1].Value; Expr = $seg.Substring($m.Length) })
    }
    return , $items
}

function Join-EaPov {
    param([object[]]$Items)
    return (@($Items | ForEach-Object { $_.Dim + $_.Expr }) -join '.')
}

function Get-EaDimName {
    param([Parameter(Mandatory)][string]$Dim)
    if ($script:DimNames.ContainsKey($Dim)) { return $script:DimNames[$Dim] }
    if ($Dim -match '^C(\d+)$') { return "Custom$($Matches[1])" }
    return $Dim
}

function Get-EaDefaultPovItem {
    # Empty POV rows for a new template (S..I plus Custom1-4).
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($d in ($script:RequiredDims + @('C1', 'C2', 'C3', 'C4'))) { $items.Add([pscustomobject]@{ Dim = $d; Expr = '' }) }
    return , $items
}

function Test-EaPov {
    # Returns a list of validation messages (empty when the POV is valid).
    param([object[]]$Items)
    $errors = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($it in $Items) {
        $label = "$($it.Dim) ($(Get-EaDimName $it.Dim))"
        if ($seen.ContainsKey($it.Dim)) { $errors.Add("$label appears more than once.") }
        $seen[$it.Dim] = $true
        $e = [string]$it.Expr
        if ([string]::IsNullOrWhiteSpace($e)) { $errors.Add("$label is empty."); continue }
        if ($e[0] -ne '#' -and $e[0] -ne '{') { $errors.Add("$label must start with # (members) or { (member list), e.g. #Actual or {Top.[Hierarchy]}.") }
        if ($e.Contains(']]>')) { $errors.Add("$label must not contain ']]>'.") }
        $depth = @{ '{' = 0; '[' = 0; '<' = 0 }
        $pairs = @{ '}' = '{'; ']' = '['; '>' = '<' }
        $bad = $false
        foreach ($ch in $e.ToCharArray()) {
            $s = [string]$ch
            if ($depth.ContainsKey($s)) { $depth[$s]++ }
            elseif ($pairs.ContainsKey($s)) { $depth[$pairs[$s]]--; if ($depth[$pairs[$s]] -lt 0) { $bad = $true } }
        }
        if ($bad -or $depth['{'] -ne 0 -or $depth['['] -ne 0 -or $depth['<'] -ne 0) { $errors.Add("$label has unbalanced brackets.") }
        elseif ([regex]::IsMatch($e, '\.' + $script:DimPattern)) {
            # Would be split as a new dimension when HFM (and this tool) re-read the POV.
            $outside = $true
            try { $outside = (Split-EaPov ($it.Dim + $e)).Count -gt 1 } catch { }
            if ($outside) { $errors.Add("$label contains '.<dim>#' or '.<dim>{' outside brackets, which would be read as another dimension.") }
        }
    }
    foreach ($d in $script:RequiredDims) {
        if (-not $seen.ContainsKey($d)) { $errors.Add("$d ($(Get-EaDimName $d)) is missing.") }
    }
    return , $errors
}

#endregion

#region Template operations

function Read-EaTemplate {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Key)
    return Read-UserParam -App $App -User $User -Key $Key
}

function Write-EaTemplateCore {
    # Writes SS_PREF_<name> and makes sure the index lists it. Caller owns the transaction.
    param([string]$App, [string]$User, [string]$Name, [string]$Xml, $Transaction)
    $err = Test-EaTemplateName $Name
    if ($err) { throw $err }
    [void](ConvertFrom-EaTemplateXml $Xml)
    $key = Get-EaTemplateKey $Name
    Write-UserParam -App $App -User $User -Key $key -Text $Xml -Transaction $Transaction
    $index = Get-EaIndex -App $App -User $User -Transaction $Transaction
    if (@($index | Where-Object { $_.Key -eq $key }).Count -eq 0) {
        $index.Add([pscustomobject]@{ Name = $Name; Key = $key })
        Set-EaIndex -App $App -User $User -Entries $index.ToArray() -Transaction $Transaction
    }
}

function Test-EaTemplateExists {
    param([string]$App, [string]$User, [string]$Name, $Transaction)
    $key = Get-EaTemplateKey $Name
    if ($null -ne (Read-UserParamBytes -App $App -User $User -Key $key -Transaction $Transaction)) { return $true }
    $index = Get-EaIndex -App $App -User $User -Transaction $Transaction
    return (@($index | Where-Object { $_.Key -eq $key }).Count -gt 0)
}

function Save-EaTemplate {
    <# Creates or overwrites a template. With -New, fails if the name already exists. #>
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Xml, [switch]$New)
    $err = Test-EaTemplateName $Name
    if ($err) { throw $err }
    $key = Get-EaTemplateKey $Name
    $backup = Backup-EaRows -App $App -Targets @([pscustomobject]@{ User = $User; Key = $script:IndexKey }, [pscustomobject]@{ User = $User; Key = $key }) -Tag "save_$Name"
    Invoke-EaTransaction {
        param($tx)
        if ($New -and (Test-EaTemplateExists -App $App -User $User -Name $Name -Transaction $tx)) { throw "A template named '$Name' already exists for $User." }
        Write-EaTemplateCore -App $App -User $User -Name $Name -Xml $Xml -Transaction $tx
    }
    return $backup
}

function Rename-EaTemplate {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$OldName, [Parameter(Mandatory)][string]$NewName)
    $err = Test-EaTemplateName $NewName
    if ($err) { throw $err }
    if ($OldName -ceq $NewName) { return }
    $oldKey = Get-EaTemplateKey $OldName
    $newKey = Get-EaTemplateKey $NewName
    $caseOnly = $oldKey -eq $newKey
    $backup = Backup-EaRows -App $App -Targets @(
        [pscustomobject]@{ User = $User; Key = $script:IndexKey },
        [pscustomobject]@{ User = $User; Key = $oldKey },
        [pscustomobject]@{ User = $User; Key = $newKey }) -Tag "rename_$OldName"
    Invoke-EaTransaction {
        param($tx)
        if (-not $caseOnly -and (Test-EaTemplateExists -App $App -User $User -Name $NewName -Transaction $tx)) { throw "A template named '$NewName' already exists for $User." }
        $bytes = Read-UserParamBytes -App $App -User $User -Key $oldKey -Transaction $tx
        if ($null -eq $bytes) { throw "Template '$OldName' has no data to rename." }
        Remove-UserParam -App $App -User $User -Key $oldKey -Transaction $tx
        Write-UserParamBytes -App $App -User $User -Key $newKey -Bytes $bytes -Transaction $tx
        $index = Get-EaIndex -App $App -User $User -Transaction $tx
        $entries = New-Object System.Collections.Generic.List[object]
        $replaced = $false
        foreach ($e in $index) {
            if ($e.Key -eq $oldKey -or $e.Key -eq $newKey) {
                if (-not $replaced) { $entries.Add([pscustomobject]@{ Name = $NewName; Key = $newKey }); $replaced = $true }
            }
            else { $entries.Add($e) }
        }
        if (-not $replaced) { $entries.Add([pscustomobject]@{ Name = $NewName; Key = $newKey }) }
        Set-EaIndex -App $App -User $User -Entries $entries.ToArray() -Transaction $tx
    }
    return $backup
}

function Remove-EaTemplate {
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string[]]$Keys)
    $targets = @([pscustomobject]@{ User = $User; Key = $script:IndexKey }) + @($Keys | ForEach-Object { [pscustomobject]@{ User = $User; Key = $_ } })
    $backup = Backup-EaRows -App $App -Targets $targets -Tag 'delete'
    Invoke-EaTransaction {
        param($tx)
        foreach ($k in $Keys) {
            if (-not $k.StartsWith($script:KeyPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Refusing to delete non-template key '$k'." }
            Remove-UserParam -App $App -User $User -Key $k -Transaction $tx
        }
        $index = Get-EaIndex -App $App -User $User -Transaction $tx
        $keep = @($index | Where-Object { $Keys -notcontains $_.Key })
        if ($keep.Count -ne $index.Count) { Set-EaIndex -App $App -User $User -Entries $keep -Transaction $tx }
    }
    return $backup
}

function Get-UniqueEaTemplateName {
    param([string]$App, [string]$User, [string]$Name, [string[]]$Reserved, $Transaction)
    $max = $script:MaxKeyLength - $script:KeyPrefix.Length
    for ($n = 1; $n -lt 1000; $n++) {
        $suffix = if ($n -eq 1) { '_copy' } else { "_copy$n" }
        $base = if ($Name.Length + $suffix.Length -gt $max) { $Name.Substring(0, $max - $suffix.Length) } else { $Name }
        $candidate = $base + $suffix
        if ($Reserved -contains $candidate) { continue }
        if (-not (Test-EaTemplateExists -App $App -User $User -Name $candidate -Transaction $Transaction)) { return $candidate }
    }
    throw "Could not find a free name for '$Name'."
}

function Add-EaTemplateSet {
    <#
        Writes a set of templates (objects with Name and Xml) to one or more users of an application
        in a single transaction. OnConflict: Overwrite | Skip | Rename (adds _copy, _copy2, ...).
        Returns one result object per template per user.
    #>
    param(
        [Parameter(Mandatory)][string]$App,
        [Parameter(Mandatory)][string[]]$Users,
        [Parameter(Mandatory)][object[]]$Templates,
        [ValidateSet('Overwrite', 'Skip', 'Rename')][string]$OnConflict = 'Skip',
        [string]$Tag = 'add'
    )
    foreach ($tpl in $Templates) {
        $err = Test-EaTemplateName $tpl.Name
        if ($err) { throw "$($tpl.Name): $err" }
        [void](ConvertFrom-EaTemplateXml $tpl.Xml)
    }
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($u in $Users) {
        $targets.Add([pscustomobject]@{ User = $u; Key = $script:IndexKey })
        foreach ($tpl in $Templates) { $targets.Add([pscustomobject]@{ User = $u; Key = (Get-EaTemplateKey $tpl.Name) }) }
    }
    $backup = Backup-EaRows -App $App -Targets $targets.ToArray() -Tag $Tag
    $results = Invoke-EaTransaction {
        param($tx)
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($u in $Users) {
            $written = New-Object System.Collections.Generic.List[string]
            foreach ($tpl in $Templates) {
                $name = $tpl.Name
                $action = 'Created'
                if ($written -contains $name -or (Test-EaTemplateExists -App $App -User $u -Name $name -Transaction $tx)) {
                    switch ($OnConflict) {
                        'Skip' { $action = 'Skipped' }
                        'Overwrite' { $action = 'Overwritten' }
                        'Rename' { $name = Get-UniqueEaTemplateName -App $App -User $u -Name $name -Reserved $written.ToArray() -Transaction $tx; $action = 'Created as copy' }
                    }
                }
                if ($action -ne 'Skipped') {
                    Write-EaTemplateCore -App $App -User $u -Name $name -Xml $tpl.Xml -Transaction $tx
                    $written.Add($name)
                }
                $out.Add([pscustomobject]@{ App = $App; User = $u; Source = $tpl.Name; Name = $name; Action = $action })
            }
        }
        , $out
    }
    return [pscustomobject]@{ Results = @($results); Backup = $backup }
}

function Copy-EaTemplate {
    param(
        [Parameter(Mandatory)][string]$SourceApp,
        [Parameter(Mandatory)][string]$SourceUser,
        [Parameter(Mandatory)][string[]]$Keys,
        [Parameter(Mandatory)][string]$TargetApp,
        [Parameter(Mandatory)][string[]]$TargetUsers,
        [ValidateSet('Overwrite', 'Skip', 'Rename')][string]$OnConflict = 'Skip'
    )
    $templates = foreach ($k in $Keys) {
        $xml = Read-UserParam -App $SourceApp -User $SourceUser -Key $k
        if ($null -eq $xml) { throw "Template '$k' has no data." }
        $name = if ($k -eq $script:DefaultKey) { 'DEFAULT' } else { $k.Substring($script:KeyPrefix.Length) }
        [pscustomobject]@{ Name = $name; Xml = $xml }
    }
    return Add-EaTemplateSet -App $TargetApp -Users $TargetUsers -Templates @($templates) -OnConflict $OnConflict -Tag 'copy'
}

function Export-EaTemplate {
    <# Writes templates to a UTF-8 XML file. The raw povTemplate XML is stored as escaped text. #>
    param([Parameter(Mandatory)][string]$App, [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string[]]$Keys, [Parameter(Mandatory)][string]$Path)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Entitize
    $w = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $w.WriteStartElement('EATemplates')
        $w.WriteAttributeString('version', '1')
        $w.WriteAttributeString('app', $App)
        $w.WriteAttributeString('user', $User)
        $w.WriteAttributeString('exported', (Get-Date).ToString('s'))
        foreach ($k in $Keys) {
            $xml = Read-UserParam -App $App -User $User -Key $k
            if ($null -eq $xml) { throw "Template '$k' has no data." }
            $name = if ($k -eq $script:DefaultKey) { 'DEFAULT' } else { $k.Substring($script:KeyPrefix.Length) }
            $w.WriteStartElement('EATemplate')
            $w.WriteAttributeString('name', $name)
            $w.WriteString($xml)
            $w.WriteEndElement()
        }
        $w.WriteEndElement()
    }
    finally { $w.Dispose() }
}

function Import-EaTemplateFile {
    <# Reads an export file and returns objects with Name and Xml (nothing is written to the database). #>
    param([Parameter(Mandatory)][string]$Path)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.Load($Path)
    if ($doc.DocumentElement.Name -ne 'EATemplates') { throw "'$Path' is not an EA template export file." }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($n in $doc.SelectNodes('/EATemplates/EATemplate')) {
        $xml = $n.InnerText
        [void](ConvertFrom-EaTemplateXml $xml)
        $list.Add([pscustomobject]@{ Name = $n.GetAttribute('name'); Xml = $xml })
    }
    return , $list
}

#endregion

Export-ModuleMember -Function @(
    'Connect-EaDatabase', 'Disconnect-EaDatabase', 'Test-EaConnection',
    'Get-HfmApp', 'Get-EaDsn', 'Get-EaTemplateUser', 'Get-HfmUser', 'Get-HfmActiveUser',
    'Read-UserParam', 'Read-UserParamBytes',
    'Backup-EaRows', 'Restore-EaBackup',
    'ConvertFrom-EaIndexXml', 'ConvertTo-EaIndexXml', 'Get-EaIndex', 'Get-EaTemplateList', 'Repair-EaIndex',
    'Get-EaTemplateKey', 'Test-EaTemplateName', 'Get-NewEaTemplateXml',
    'ConvertFrom-EaTemplateXml', 'ConvertTo-EaTemplateXml',
    'Split-EaPov', 'Join-EaPov', 'Get-EaDimName', 'Get-EaDefaultPovItem', 'Test-EaPov',
    'Read-EaTemplate', 'Save-EaTemplate', 'Rename-EaTemplate', 'Remove-EaTemplate',
    'Add-EaTemplateSet', 'Copy-EaTemplate', 'Export-EaTemplate', 'Import-EaTemplateFile'
)

}
#endregion

#region Embedded window layouts (source: Xaml\*.xaml)
$script:XamlText = @{
    'ConnectDialog.xaml' = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Connect to HFM database" Width="440" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" FontSize="13" ShowInTaskbar="False">
    <Grid Margin="14">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="SQL Server" VerticalAlignment="Center" Margin="0,4,10,4"/>
        <TextBox Grid.Row="0" Grid.Column="1" x:Name="txtServer" Margin="0,4" ToolTip="server, server\instance or server,port"/>
        <TextBlock Grid.Row="1" Text="Database" VerticalAlignment="Center" Margin="0,4,10,4"/>
        <TextBox Grid.Row="1" Grid.Column="1" x:Name="txtDatabase" Margin="0,4"/>
        <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="0,8,0,4">
            <RadioButton x:Name="rbWindows" Content="Windows authentication" Margin="0,0,16,0" IsChecked="True"/>
            <RadioButton x:Name="rbSql" Content="SQL login"/>
        </StackPanel>
        <TextBlock Grid.Row="3" Text="Login" VerticalAlignment="Center" Margin="0,4,10,4"/>
        <TextBox Grid.Row="3" Grid.Column="1" x:Name="txtSqlUser" Margin="0,4"/>
        <TextBlock Grid.Row="4" Text="Password" VerticalAlignment="Center" Margin="0,4,10,4"/>
        <PasswordBox Grid.Row="4" Grid.Column="1" x:Name="pwdSql" Margin="0,4"/>
        <TextBlock Grid.Row="5" Grid.ColumnSpan="2" x:Name="txtError" Foreground="#FFC62828" TextWrapping="Wrap" Margin="0,8,0,0"/>
        <StackPanel Grid.Row="6" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOK" Content="Connect" IsDefault="True" Padding="14,3" Margin="0,0,8,0"/>
            <Button x:Name="btnCancel" Content="Cancel" IsCancel="True" Padding="14,3"/>
        </StackPanel>
    </Grid>
</Window>
'@
    'CopyDialog.xaml' = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Copy templates" Width="520" Height="560" MinHeight="420"
        WindowStartupLocation="CenterOwner" FontSize="13" ShowInTaskbar="False">
    <DockPanel Margin="14">
        <TextBlock DockPanel.Dock="Top" x:Name="txtHeader" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Target application" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <ComboBox x:Name="cmbTargetApp" Width="200"/>
        </StackPanel>
        <TextBlock DockPanel.Dock="Top" x:Name="txtAppWarn" Visibility="Collapsed" Foreground="#FFA35200" TextWrapping="Wrap" Margin="0,0,0,8"
                   Text="The target is a different application: check that the POV members and the DSN are valid there."/>
        <StackPanel DockPanel.Dock="Bottom">
            <TextBlock Text="If a template with the same name exists" FontWeight="SemiBold" Margin="0,10,0,4"/>
            <RadioButton x:Name="rbSkip" Content="Skip it" IsChecked="True" Margin="0,2"/>
            <RadioButton x:Name="rbOverwrite" Content="Overwrite it" Margin="0,2"/>
            <RadioButton x:Name="rbRename" Content="Keep both (new one gets a _copy suffix)" Margin="0,2"/>
            <TextBlock x:Name="txtError" Foreground="#FFC62828" TextWrapping="Wrap" Margin="0,8,0,0"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
                <Button x:Name="btnOK" Content="Copy" IsDefault="True" Padding="14,3" Margin="0,0,8,0" MinWidth="72"/>
                <Button x:Name="btnCancel" Content="Cancel" IsCancel="True" Padding="14,3" MinWidth="72"/>
            </StackPanel>
        </StackPanel>
        <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
            <Button DockPanel.Dock="Right" x:Name="btnAddUser" Content="Add" Padding="10,3" Margin="6,0,0,0"/>
            <TextBox x:Name="txtNewUser" ToolTip="Type a user name exactly as HFM stores it, e.g. jsmith@Native Directory"/>
        </DockPanel>
        <TextBlock DockPanel.Dock="Top" Text="Target users (click to select one or more)" FontWeight="SemiBold" Margin="0,0,0,4"/>
        <ListBox x:Name="lstUsers" SelectionMode="Multiple"/>
    </DockPanel>
</Window>
'@
    'MainWindow.xaml' = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="HFM EA Template Manager" Width="1220" Height="780" MinWidth="900" MinHeight="560"
        WindowStartupLocation="CenterScreen" FontSize="13">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,3"/>
            <Setter Property="Margin" Value="0,0,6,6"/>
            <Setter Property="MinWidth" Value="72"/>
        </Style>
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,4,10,4"/>
        </Style>
        <Style x:Key="Mono" TargetType="Control">
            <Setter Property="FontFamily" Value="Consolas"/>
        </Style>
    </Window.Resources>
    <DockPanel>
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem x:Name="mnuConnect" Header="_Connect..."/>
                <Separator/>
                <MenuItem x:Name="mnuBackupAll" Header="_Back up all EA rows of this application"/>
                <MenuItem x:Name="mnuRestore" Header="_Restore from backup..."/>
                <MenuItem x:Name="mnuOpenBackups" Header="Open backup _folder"/>
                <Separator/>
                <MenuItem x:Name="mnuExit" Header="E_xit"/>
            </MenuItem>
            <MenuItem Header="_Tools">
                <MenuItem x:Name="mnuRefresh" Header="_Refresh" InputGestureText="F5"/>
                <MenuItem x:Name="mnuRepairIndex" Header="Repair template _index for this user"/>
            </MenuItem>
        </Menu>

        <StatusBar DockPanel.Dock="Bottom">
            <TextBlock x:Name="txtStatus" Text="Not connected."/>
        </StatusBar>

        <Border DockPanel.Dock="Top" Padding="10,8" BorderThickness="0,0,0,1" BorderBrush="#FFD0D0D0" Background="#FFF5F6F8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Grid.Column="0" Text="Application" Style="{StaticResource FieldLabel}"/>
                <ComboBox Grid.Column="1" x:Name="cmbApp" Width="170" Margin="0,0,16,0"/>
                <TextBlock Grid.Column="2" Text="User" Style="{StaticResource FieldLabel}"/>
                <ComboBox Grid.Column="3" x:Name="cmbUser" Width="300" IsEditable="True"
                          ToolTip="Pick a user with templates, or type a user name (user@Provider) and press Enter."/>
                <Button Grid.Column="4" x:Name="btnLoadUser" Content="Load" Margin="6,0,0,0"/>
                <TextBlock Grid.Column="6" x:Name="txtConnection" Text="Not connected" VerticalAlignment="Center" Foreground="#FF555555"/>
                <TextBlock Grid.Row="1" Grid.ColumnSpan="7" x:Name="txtActiveWarn" Visibility="Collapsed" Margin="0,6,0,0"
                           Foreground="#FFA35200" TextWrapping="Wrap"
                           Text="This user has an active HFM session on this application. Ask them to log off first: HFM may not see these changes, or may overwrite them, until they log in again."/>
            </Grid>
        </Border>

        <Grid Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300" MinWidth="220"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*" MinWidth="420"/>
            </Grid.ColumnDefinitions>

            <!-- Template list -->
            <DockPanel Grid.Column="0">
                <TextBlock DockPanel.Dock="Top" Text="Templates" FontWeight="SemiBold" Margin="0,0,0,6"/>
                <WrapPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
                    <Button x:Name="btnNew" Content="New..."/>
                    <Button x:Name="btnRename" Content="Rename..."/>
                    <Button x:Name="btnDelete" Content="Delete"/>
                    <Button x:Name="btnCopy" Content="Copy to..."/>
                    <Button x:Name="btnExport" Content="Export..."/>
                    <Button x:Name="btnImport" Content="Import..."/>
                </WrapPanel>
                <ListBox x:Name="lstTemplates" SelectionMode="Extended"
                         ToolTip="Ctrl/Shift-click to select several templates for Copy, Export or Delete."/>
            </DockPanel>

            <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch" Background="Transparent"/>

            <!-- Editor -->
            <DockPanel Grid.Column="2">
                <StackPanel DockPanel.Dock="Top" Margin="0,0,0,6">
                    <TextBlock x:Name="txtTemplateTitle" Text="No template selected" FontSize="16" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtTemplateInfo" Foreground="#FF555555" TextWrapping="Wrap"/>
                </StackPanel>
                <WrapPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
                    <Button x:Name="btnSave" Content="Save" FontWeight="SemiBold"/>
                    <Button x:Name="btnRevert" Content="Revert"/>
                    <Button x:Name="btnSaveAs" Content="Save as new..."/>
                </WrapPanel>
                <TabControl x:Name="tabEditor" IsEnabled="False">
                    <TabItem Header="POV">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="pnlPovTab" Margin="10">
                                <TextBlock TextWrapping="Wrap" Foreground="#FF555555" Margin="0,0,0,8"
                                           Text="Each expression starts with # for members (e.g. #Jan;Feb;Mar) or { for member lists (e.g. {Top.[Hierarchy]}). Combine several with ; (e.g. {[Base]};{[Hierarchy]})."/>
                                <Grid x:Name="gridPov">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="40"/>
                                        <ColumnDefinition Width="90"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                </Grid>
                                <WrapPanel Margin="0,8,0,0">
                                    <Button x:Name="btnAddCustom" Content="+ Custom"/>
                                    <Button x:Name="btnRemoveCustom" Content="- Custom"/>
                                </WrapPanel>
                                <TextBlock x:Name="txtPovErrors" Foreground="#FFC62828" TextWrapping="Wrap" Margin="0,4,0,8"/>
                                <TextBlock Text="Resulting POV" FontWeight="SemiBold" Margin="0,4,0,4"/>
                                <TextBox x:Name="txtPovFull" IsReadOnly="True" TextWrapping="Wrap" Style="{StaticResource Mono}" Background="#FFF7F7F7"/>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>
                    <TabItem Header="Options">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <Grid x:Name="gridOptions" Margin="10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="320"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Text="Extract format" Style="{StaticResource FieldLabel}"/>
                                <ComboBox Grid.Row="0" Grid.Column="1" x:Name="cmbExtractFormat" IsEditable="True" Margin="0,3"
                                          ToolTip="extractFormat code. Labels come from EATemplateManager.config.json."/>
                                <TextBlock Grid.Row="1" Text="Destination DSN" Style="{StaticResource FieldLabel}"/>
                                <ComboBox Grid.Row="1" Grid.Column="1" x:Name="cmbDsn" IsEditable="True" Margin="0,3" ToolTip="selectedDSN"/>
                                <TextBlock Grid.Row="2" Text="Table prefix" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="2" Grid.Column="1" x:Name="txtTablePrefix" Margin="0,3" ToolTip="tablePrefix"/>
                                <TextBlock Grid.Row="3" Text="File prefix" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="3" Grid.Column="1" x:Name="txtFilePrefix" Margin="0,3" ToolTip="filePrefix"/>
                                <TextBlock Grid.Row="4" Text="Delimiter" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="4" Grid.Column="1" x:Name="txtDelimiter" Margin="0,3" Width="60" HorizontalAlignment="Left" ToolTip="delimiter"/>
                                <TextBlock Grid.Row="5" Text="Export option" Style="{StaticResource FieldLabel}"/>
                                <TextBox Grid.Row="5" Grid.Column="1" x:Name="txtExportOption" Margin="0,3" Width="60" HorizontalAlignment="Left" ToolTip="exportOption (numeric code)"/>
                                <TextBlock Grid.Row="6" Text="Line item detail" Style="{StaticResource FieldLabel}"/>
                                <ComboBox Grid.Row="6" Grid.Column="1" x:Name="cmbLineItem" IsEditable="True" Margin="0,3" ToolTip="lineItemDetailOption"/>
                                <TextBlock Grid.Row="7" Text="Extract" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,8,10,4"/>
                                <StackPanel Grid.Row="7" Grid.Column="1" Margin="0,6,0,0">
                                    <CheckBox x:Name="chkExtractData" Content="Data (extractData)" Margin="0,2"/>
                                    <CheckBox x:Name="chkExtractCalculatedData" Content="Calculated data (extractCalculatedData)" Margin="0,2"/>
                                    <CheckBox x:Name="chkExtractDerivedData" Content="Derived data (extractDerivedData)" Margin="0,2"/>
                                    <CheckBox x:Name="chkExtractDynAccts" Content="Dynamic accounts (extractDynAccts)" Margin="0,2"/>
                                    <CheckBox x:Name="chkExtractPhasedSubmission" Content="Phased submission (extractPhasedSubmission)" Margin="0,2"/>
                                    <CheckBox x:Name="chkExtractCellText" Content="Cell text (extractCellText)" Margin="0,2"/>
                                </StackPanel>
                                <TextBlock Grid.Row="8" Text="Other options" Style="{StaticResource FieldLabel}" VerticalAlignment="Top" Margin="0,10,10,4"/>
                                <TextBlock Grid.Row="8" Grid.Column="1" Grid.ColumnSpan="2" x:Name="txtOtherOptions" Margin="0,10,0,0" TextWrapping="Wrap" Foreground="#FF555555"
                                           Text="(none)"/>
                            </Grid>
                        </ScrollViewer>
                    </TabItem>
                    <TabItem x:Name="tabRaw" Header="Raw XML">
                        <DockPanel Margin="10">
                            <CheckBox DockPanel.Dock="Top" x:Name="chkEditRaw" Margin="0,0,0,8"
                                      Content="Edit raw XML (advanced: the POV and Options tabs are locked while this is on)"/>
                            <TextBox x:Name="txtRaw" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                     VerticalScrollBarVisibility="Auto" Style="{StaticResource Mono}"/>
                        </DockPanel>
                    </TabItem>
                </TabControl>
            </DockPanel>
        </Grid>
    </DockPanel>
</Window>
'@
    'NameDialog.xaml' = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Template name" Width="420" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" FontSize="13" ShowInTaskbar="False">
    <StackPanel Margin="14">
        <TextBlock x:Name="txtPrompt" Text="Template name:" TextWrapping="Wrap" Margin="0,0,0,6"/>
        <TextBox x:Name="txtName"/>
        <TextBlock x:Name="txtError" Foreground="#FFC62828" TextWrapping="Wrap" Margin="0,8,0,0"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOK" Content="OK" IsDefault="True" Padding="14,3" Margin="0,0,8,0" MinWidth="72"/>
            <Button x:Name="btnCancel" Content="Cancel" IsCancel="True" Padding="14,3" MinWidth="72"/>
        </StackPanel>
    </StackPanel>
</Window>
'@
}
#endregion

$script:Root = $PSScriptRoot
$script:ConfigPath = Join-Path $PSScriptRoot 'EATemplateManager.config.json'
$script:PrefsPath = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'EATemplateManager\preferences.json'
$script:PrefKeys = @('Server', 'Database', 'UseSqlAuth', 'SqlUser', 'LastApp', 'LastUser')
$script:BoolOptions = [ordered]@{
    extractDynAccts         = 'chkExtractDynAccts'
    extractData             = 'chkExtractData'
    extractCalculatedData   = 'chkExtractCalculatedData'
    extractDerivedData      = 'chkExtractDerivedData'
    extractPhasedSubmission = 'chkExtractPhasedSubmission'
    extractCellText         = 'chkExtractCellText'
}
$script:TextOptions = [ordered]@{
    tablePrefix          = 'txtTablePrefix'
    filePrefix           = 'txtFilePrefix'
    exportOption         = 'txtExportOption'
    extractFormat        = 'cmbExtractFormat'
    selectedDSN          = 'cmbDsn'
    lineItemDetailOption = 'cmbLineItem'
    delimiter            = 'txtDelimiter'
}

# Session state
$script:S = @{
    App               = $null
    User              = $null
    Current           = $null    # @{ Name; Key; Status; IsDefault; IsNew }
    CurrentItem       = $null    # ListBoxItem of the current template
    BaseXml           = $null
    Dirty             = $false
    Loading           = $false
    RawMode           = $false
    SuppressSelection = $false
    SuppressApp       = $false
    PovRows           = $null
    ActiveUsers       = @()
}

#region Helpers

function New-XamlWindow {
    param([string]$File)
    $text = $script:XamlText[$File]
    $window = [Windows.Markup.XamlReader]::Parse($text)
    $controls = @{}
    foreach ($m in [regex]::Matches($text, 'x:Name="([^"]+)"')) {
        $controls[$m.Groups[1].Value] = $window.FindName($m.Groups[1].Value)
    }
    return @{ Window = $window; C = $controls }
}

function Read-JsonFile([string]$Path) {
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Raw -Encoding UTF8 $Path | ConvertFrom-Json) } catch { return $null }
}

function Get-Config {
    <#
        Settings are layered, later layers win:
          1. built-in defaults
          2. EATemplateManager.config.json next to the script (optional, shared: label overrides etc.)
          3. %APPDATA%\EATemplateManager\preferences.json (per Windows user: connection and last selection)
        Passwords are never stored.
    #>
    # Extract format order confirmed from HFM's xfm.en-us.csv resources.
    $formats = [pscustomobject][ordered]@{
        '0' = 'Standard'; '1' = 'Metadata Only'; '2' = 'Selected Metadata Only'; '3' = 'SQL Aggregation'
        '4' = 'Essbase'; '5' = 'Data Warehouse'; '6' = 'Flat File'; '7' = 'Flat File without Header'
    }
    $cfg = [pscustomobject]@{
        Server = 'localhost'; Database = 'hfm'; UseSqlAuth = $false; SqlUser = ''; LastApp = ''; LastUser = ''
        ExtractFormats = $formats; LineItemOptions = @('EA_LINEITEM_EXCLUDE', 'EA_LINEITEM_SUMMARY', 'EA_LINEITEM_DETAIL')
    }
    $shared = Read-JsonFile $script:ConfigPath
    if ($shared) {
        foreach ($p in $shared.PSObject.Properties) { $cfg | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force }
    }
    $prefs = Read-JsonFile $script:PrefsPath
    if ($prefs) {
        foreach ($k in $script:PrefKeys) {
            $p = $prefs.PSObject.Properties[$k]
            if ($p) { $cfg.$k = $p.Value }
        }
    }
    return $cfg
}

function Save-Config {
    # Writes only the per-user preferences; never a password.
    $prefs = [ordered]@{}
    foreach ($k in $script:PrefKeys) { $prefs[$k] = $script:Config.$k }
    try {
        $dir = Split-Path -Parent $script:PrefsPath
        if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
        [pscustomobject]$prefs | ConvertTo-Json | Set-Content -Path $script:PrefsPath -Encoding UTF8
    }
    catch {
        if (-not $script:PrefsWarned) {
            $script:PrefsWarned = $true
            Set-Status "Could not save preferences to $($script:PrefsPath): $($_.Exception.Message)"
        }
    }
}

function Set-Status([string]$Text) {
    $script:C.txtStatus.Text = $Text
}

function Show-Message([string]$Text, [string]$Title = 'EA Template Manager', [string]$Icon = 'Information') {
    [void][System.Windows.MessageBox]::Show($script:Window, $Text, $Title, 'OK', $Icon)
}

function Confirm-Action([string]$Text, [string]$Title = 'Confirm') {
    return ([System.Windows.MessageBox]::Show($script:Window, $Text, $Title, 'YesNo', 'Question') -eq 'Yes')
}

function Invoke-Safe {
    # Runs a UI action; shows errors instead of letting them escape the event handler.
    param([scriptblock]$Action)
    try { & $Action }
    catch {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        try { "[$(Get-Date -Format s)] $($_.Exception.GetType().FullName): $($_.Exception.Message)`r`n$($_.ScriptStackTrace)`r`n" | Add-Content -Path (Join-Path $script:Root 'EATemplateManager.log') -Encoding UTF8 } catch { }
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException -and $_.Exception.InnerException.Message -ne $msg) { $msg += "`n`n" + $_.Exception.InnerException.Message }
        Set-Status "Error: $msg"
        Show-Message $msg 'Error' 'Error'
    }
}

function Get-ExtractFormatText([string]$Code) {
    $p = $script:Config.ExtractFormats.PSObject.Properties[$Code]
    if ($p) { return "$Code - $($p.Value)" }
    return $Code
}

function Test-UserActive {
    return [bool]($script:S.User -and (@($script:S.ActiveUsers | Where-Object { $_ -eq $script:S.User }).Count -gt 0))
}

#endregion

#region Connection, application and user

function Show-ConnectDialog {
    $d = New-XamlWindow 'ConnectDialog.xaml'
    $script:Dlg = $d
    $d.Window.Owner = $script:Window
    $d.C.txtServer.Text = $script:Config.Server
    $d.C.txtDatabase.Text = $script:Config.Database
    $d.C.txtSqlUser.Text = $script:Config.SqlUser
    $d.C.rbSql.IsChecked = [bool]$script:Config.UseSqlAuth
    $d.C.rbWindows.IsChecked = -not [bool]$script:Config.UseSqlAuth
    $toggle = {
        $sql = [bool]$script:Dlg.C.rbSql.IsChecked
        $script:Dlg.C.txtSqlUser.IsEnabled = $sql
        $script:Dlg.C.pwdSql.IsEnabled = $sql
    }
    & $toggle
    $d.C.rbSql.Add_Checked($toggle)
    $d.C.rbWindows.Add_Checked($toggle)
    $d.C.btnOK.Add_Click({
            $c = $script:Dlg.C
            $c.txtError.Text = ''
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try {
                $connArgs = @{ Server = $c.txtServer.Text.Trim(); Database = $c.txtDatabase.Text.Trim() }
                if ($c.rbSql.IsChecked) {
                    if (-not $c.txtSqlUser.Text.Trim()) { throw 'Enter the SQL login name.' }
                    $connArgs.UserName = $c.txtSqlUser.Text.Trim()
                    $connArgs.Password = $c.pwdSql.SecurePassword
                }
                Connect-EaDatabase @connArgs
                $script:Config.Server = $connArgs.Server
                $script:Config.Database = $connArgs.Database
                $script:Config.UseSqlAuth = [bool]$c.rbSql.IsChecked
                $script:Config.SqlUser = $c.txtSqlUser.Text.Trim()
                Save-Config
                $script:Dlg.Window.DialogResult = $true
            }
            catch { $c.txtError.Text = $_.Exception.Message }
            finally { [System.Windows.Input.Mouse]::OverrideCursor = $null }
        })
    $d.Window.Add_Loaded({ if ($script:Config.UseSqlAuth) { [void]$script:Dlg.C.pwdSql.Focus() } else { [void]$script:Dlg.C.txtServer.Focus() } })
    if ($d.Window.ShowDialog()) {
        $how = if ($script:Config.UseSqlAuth) { "SQL login $($script:Config.SqlUser)" } else { 'Windows auth' }
        $script:C.txtConnection.Text = "$($script:Config.Server) / $($script:Config.Database)  ($how)"
        Initialize-AfterConnect
    }
}

function Initialize-AfterConnect {
    $dsns = @(Get-EaDsn)
    $script:C.cmbDsn.ItemsSource = $dsns
    $apps = @(Get-HfmApp)
    $script:S.SuppressApp = $true
    $script:C.cmbApp.ItemsSource = $apps
    $script:S.SuppressApp = $false
    Set-Status "Connected. $($apps.Count) application(s), $($dsns.Count) EA DSN(s)."
    if ($apps.Count -eq 0) { Show-Message 'No HFM applications with a USERPARAMS table were found in this database.'; return }
    $pick = if ($apps -contains $script:Config.LastApp) { $script:Config.LastApp } else { $apps[0] }
    $script:C.cmbApp.SelectedItem = $pick   # triggers Select-App
}

function Select-App([string]$App) {
    if (-not $App) { return }
    $script:S.App = $App
    $script:S.User = $null
    $script:Config.LastApp = $App
    Save-Config
    $script:S.ActiveUsers = @(Get-HfmActiveUser -App $App)
    $users = @(Get-EaTemplateUser -App $App)
    $script:C.cmbUser.ItemsSource = $users
    Clear-TemplateList
    $pick = $null
    if ($users -contains $script:Config.LastUser) { $pick = $script:Config.LastUser } elseif ($users.Count) { $pick = $users[0] }
    if ($pick) { $script:C.cmbUser.SelectedItem = $pick; Select-User $pick }
    else { $script:C.cmbUser.Text = ''; Set-Status "$App has no users with EA templates yet. Type a user name and press Enter." }
}

function Select-User([string]$User) {
    $User = "$User".Trim()
    if (-not $User -or -not $script:S.App) { return }
    if ($User -eq $script:S.User) { return }
    if (-not (Confirm-Discard)) {
        $script:C.cmbUser.Text = $script:S.User
        return
    }
    $script:S.User = $User
    $script:Config.LastUser = $User
    Save-Config
    $script:C.txtActiveWarn.Visibility = if (Test-UserActive) { 'Visible' } else { 'Collapsed' }
    Update-TemplateList
}

#endregion

#region Template list

function Clear-TemplateList {
    $script:S.SuppressSelection = $true
    $script:C.lstTemplates.Items.Clear()
    $script:S.SuppressSelection = $false
    Clear-Editor
    Update-Buttons
}

function Update-TemplateList([string]$SelectKey) {
    $app = $script:S.App; $user = $script:S.User
    Clear-TemplateList
    if (-not $app -or -not $user) { return }
    $entries = Get-EaTemplateList -App $app -User $user
    $hasDefault = $null -ne (Read-UserParamBytes -App $app -User $user -Key 'SS_DEFAULT')
    $items = New-Object System.Collections.Generic.List[object]
    if ($hasDefault) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = 'SS_DEFAULT  (last-used settings)'
        $tb.FontStyle = [System.Windows.FontStyles]::Italic
        $tb.Foreground = [System.Windows.Media.Brushes]::DimGray
        $li = New-Object System.Windows.Controls.ListBoxItem
        $li.Content = $tb
        $li.Tag = [pscustomobject]@{ Name = 'SS_DEFAULT'; Key = 'SS_DEFAULT'; Status = 'OK'; IsDefault = $true; IsNew = $false }
        $items.Add($li)
    }
    foreach ($e in $entries) {
        $tb = New-Object System.Windows.Controls.TextBlock
        switch ($e.Status) {
            'OK' { $tb.Text = $e.Name }
            'NotInIndex' { $tb.Text = "$($e.Name)   [not in index]"; $tb.Foreground = [System.Windows.Media.Brushes]::DarkOrange }
            'MissingData' { $tb.Text = "$($e.Name)   [no data]"; $tb.Foreground = [System.Windows.Media.Brushes]::Firebrick }
        }
        $li = New-Object System.Windows.Controls.ListBoxItem
        $li.Content = $tb
        $li.Tag = [pscustomobject]@{ Name = $e.Name; Key = $e.Key; Status = $e.Status; IsDefault = $false; IsNew = $false }
        $items.Add($li)
    }
    $script:S.SuppressSelection = $true
    foreach ($li in $items) { [void]$script:C.lstTemplates.Items.Add($li) }
    $script:S.SuppressSelection = $false
    $issues = @($entries | Where-Object Status -ne 'OK').Count
    $msg = "$app / $user : $($entries.Count) template(s)."
    if ($issues) { $msg += " $issues index issue(s): use Tools > Repair template index." }
    if (Test-UserActive) { $msg += ' User has an active HFM session.' }
    Set-Status $msg
    $select = $null
    if ($SelectKey) { $select = $items | Where-Object { $_.Tag.Key -eq $SelectKey } | Select-Object -First 1 }
    if ($select) { $script:C.lstTemplates.SelectedItem = $select }   # triggers Show-Template
    Update-Buttons
}

function Get-SelectedEntries {
    return @($script:C.lstTemplates.SelectedItems | ForEach-Object { $_.Tag })
}

function Update-Buttons {
    $c = $script:C
    $sel = @(Get-SelectedEntries)
    $templates = @($sel | Where-Object { -not $_.IsDefault })
    $withData = @($sel | Where-Object { $_.Status -ne 'MissingData' })
    $hasUser = [bool]($script:S.App -and $script:S.User)
    $cur = $script:S.Current
    $c.btnNew.IsEnabled = $hasUser
    $c.btnImport.IsEnabled = $hasUser
    $c.btnRename.IsEnabled = ($sel.Count -eq 1 -and $templates.Count -eq 1 -and $templates[0].Status -ne 'MissingData')
    $c.btnDelete.IsEnabled = ($templates.Count -gt 0 -and $templates.Count -eq $sel.Count)
    $c.btnCopy.IsEnabled = ($withData.Count -gt 0)
    $c.btnExport.IsEnabled = ($withData.Count -gt 0)
    $c.btnSave.IsEnabled = [bool]($cur -and -not $cur.IsDefault -and ($script:S.Dirty -or $cur.IsNew))
    $c.btnRevert.IsEnabled = [bool]($cur -and $script:S.Dirty -and -not $cur.IsNew)
    $c.btnSaveAs.IsEnabled = [bool]($cur -and $script:S.BaseXml)
    $c.mnuRepairIndex.IsEnabled = $hasUser
    $c.mnuBackupAll.IsEnabled = [bool]$script:S.App
    $c.mnuRestore.IsEnabled = [bool]$script:S.App
}

#endregion

#region Editor

function Clear-Editor {
    $script:S.Loading = $true
    $script:S.Current = $null
    $script:S.CurrentItem = $null
    $script:S.BaseXml = $null
    $script:S.LoadedXml = $null
    $script:S.Dirty = $false
    Set-RawMode $false
    Set-PovRows @()
    $script:C.txtRaw.Text = ''
    $script:C.txtPovFull.Text = ''
    $script:C.txtPovErrors.Text = ''
    $script:C.txtOtherOptions.Text = '(none)'
    $script:C.txtTemplateTitle.Text = 'No template selected'
    $script:C.txtTemplateInfo.Text = ''
    $script:C.tabEditor.IsEnabled = $false
    $script:S.Loading = $false
}

function Set-PovRows([object[]]$Items) {
    $g = $script:C.gridPov
    $g.Children.Clear()
    $g.RowDefinitions.Clear()
    $script:S.PovRows = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Items) { Add-PovRow $it.Dim $it.Expr }
}

function Add-PovRow([string]$Dim, [string]$Expr) {
    $g = $script:C.gridPov
    $row = $g.RowDefinitions.Count
    $rd = New-Object System.Windows.Controls.RowDefinition
    $rd.Height = [System.Windows.GridLength]::Auto
    $g.RowDefinitions.Add($rd)

    $code = New-Object System.Windows.Controls.TextBlock
    $code.Text = $Dim
    $code.FontWeight = [System.Windows.FontWeights]::Bold
    $code.FontFamily = New-Object System.Windows.Media.FontFamily 'Consolas'
    $code.VerticalAlignment = 'Center'
    $name = New-Object System.Windows.Controls.TextBlock
    $name.Text = Get-EaDimName $Dim
    $name.VerticalAlignment = 'Center'
    $name.Foreground = [System.Windows.Media.Brushes]::DimGray
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $Expr
    $tb.FontFamily = New-Object System.Windows.Media.FontFamily 'Consolas'
    $tb.Margin = New-Object System.Windows.Thickness 0, 3, 0, 3
    $tb.Add_TextChanged({ Invoke-Safe { Set-Dirty } })

    [System.Windows.Controls.Grid]::SetRow($code, $row); [System.Windows.Controls.Grid]::SetColumn($code, 0)
    [System.Windows.Controls.Grid]::SetRow($name, $row); [System.Windows.Controls.Grid]::SetColumn($name, 1)
    [System.Windows.Controls.Grid]::SetRow($tb, $row); [System.Windows.Controls.Grid]::SetColumn($tb, 2)
    [void]$g.Children.Add($code); [void]$g.Children.Add($name); [void]$g.Children.Add($tb)
    $script:S.PovRows.Add([pscustomobject]@{ Dim = $Dim; TextBox = $tb })
}

function Get-PovItems {
    return @($script:S.PovRows | ForEach-Object { [pscustomobject]@{ Dim = $_.Dim; Expr = $_.TextBox.Text.Trim() } })
}

function Set-OptionControls($Options) {
    $c = $script:C
    foreach ($k in $script:BoolOptions.Keys) {
        $c[$script:BoolOptions[$k]].IsChecked = ($Options.Contains($k) -and "$($Options[$k])" -eq 'true')
    }
    foreach ($k in $script:TextOptions.Keys) {
        $v = if ($Options.Contains($k)) { [string]$Options[$k] } else { '' }
        if ($k -eq 'extractFormat') { $v = Get-ExtractFormatText $v }
        $c[$script:TextOptions[$k]].Text = $v
    }
    $other = @($Options.Keys | Where-Object { -not $script:BoolOptions.Contains($_) -and -not $script:TextOptions.Contains($_) } | ForEach-Object { "$_ = $($Options[$_])" })
    $c.txtOtherOptions.Text = if ($other.Count) { ($other -join "`n") + "`n(kept as-is; edit in Raw XML)" } else { '(none)' }
}

function Get-OptionValues {
    $c = $script:C
    $o = [ordered]@{}
    foreach ($k in $script:TextOptions.Keys) {
        $v = $c[$script:TextOptions[$k]].Text
        if ($k -eq 'extractFormat' -and $v -match '^\s*(\d+)') { $v = $Matches[1] }
        elseif ($k -ne 'delimiter') { $v = $v.Trim() }
        $o[$k] = $v
    }
    foreach ($k in $script:BoolOptions.Keys) {
        $o[$k] = if ($c[$script:BoolOptions[$k]].IsChecked) { 'true' } else { 'false' }
    }
    return $o
}

function Get-EditorXml {
    # Current editor content as template XML (throws if the raw XML is invalid).
    if ($script:S.RawMode) {
        $xml = $script:C.txtRaw.Text
        [void](ConvertFrom-EaTemplateXml $xml)
        return $xml
    }
    $pov = Join-EaPov (Get-PovItems)
    return ConvertTo-EaTemplateXml -BaseXml $script:S.BaseXml -Pov $pov -Options (Get-OptionValues)
}

function Get-EditorErrors {
    $errors = New-Object System.Collections.Generic.List[string]
    try {
        if ($script:S.RawMode) {
            $m = ConvertFrom-EaTemplateXml $script:C.txtRaw.Text
            $items = Split-EaPov $m.Pov
            foreach ($e in (Test-EaPov $items.ToArray())) { $errors.Add($e) }
        }
        else {
            foreach ($e in (Test-EaPov (Get-PovItems))) { $errors.Add($e) }
        }
    }
    catch { $errors.Add($_.Exception.Message) }
    return , $errors
}

function Update-Validation {
    if ($script:S.RawMode) {
        $script:C.txtPovErrors.Text = ''
        return
    }
    $items = Get-PovItems
    $script:C.txtPovFull.Text = Join-EaPov $items
    $script:C.txtPovErrors.Text = ((Test-EaPov $items) -join "`n")
}

function Set-Dirty {
    # Called on every editor change: dirty means the editor content differs from what was loaded.
    if ($script:S.Loading -or -not $script:S.Current -or -not $script:S.BaseXml) { return }
    $dirty = $true
    if (-not $script:S.Current.IsNew) {
        try { $dirty = (Get-EditorXml) -cne $script:S.LoadedXml } catch { $dirty = $true }
    }
    if ($dirty -and $script:S.Current.IsDefault) {
        $script:C.txtTemplateInfo.Text = 'SS_DEFAULT is read-only here. Use "Save as new..." to keep these edits as a template.'
    }
    $script:S.Dirty = $dirty
    Update-Title
    Update-Validation
    Update-Buttons
}

function Set-RawMode([bool]$On) {
    # Switches the editor between the POV/Options form and raw XML editing (no change events).
    $was = $script:S.Loading
    $script:S.Loading = $true
    try {
        $script:C.chkEditRaw.IsChecked = $On
        $script:S.RawMode = $On
        $script:C.txtRaw.IsReadOnly = -not $On
        $script:C.pnlPovTab.IsEnabled = -not $On
        $script:C.gridOptions.IsEnabled = -not $On
    }
    finally { $script:S.Loading = $was }
}

function Update-Title {
    $cur = $script:S.Current
    if (-not $cur) { return }
    $star = if ($script:S.Dirty) { ' *' } else { '' }
    if ($cur.IsDefault) { $script:C.txtTemplateTitle.Text = "SS_DEFAULT$star" }
    elseif ($cur.IsNew) { $script:C.txtTemplateTitle.Text = "$($cur.Name)  (new, not saved)$star" }
    else { $script:C.txtTemplateTitle.Text = "$($cur.Name)$star" }
}

function Show-Template($Entry, $Xml) {
    # Loads a template into the editor. $Xml overrides reading from the database (used for new templates).
    $script:S.Loading = $true
    try {
        Set-RawMode $false
        $script:S.Current = $Entry
        $script:S.Dirty = [bool]$Entry.IsNew
        if ($null -eq $Xml) { $Xml = Read-EaTemplate -App $script:S.App -User $script:S.User -Key $Entry.Key }
        if ($null -eq $Xml) {
            $script:S.BaseXml = $null
            $script:S.LoadedXml = $null
            Set-PovRows @()
            $script:C.txtRaw.Text = ''
            $script:C.tabEditor.IsEnabled = $false
            $script:C.txtTemplateTitle.Text = $Entry.Name
            $script:C.txtTemplateInfo.Text = "The index lists $($Entry.Key) but it has no data. Delete it, or use Tools > Repair template index."
            return
        }
        $script:S.BaseXml = $Xml
        $model = ConvertFrom-EaTemplateXml $Xml
        $items = $null
        try { $items = Split-EaPov $model.Pov } catch { }
        if ($Entry.IsNew -and ($null -eq $items -or $items.Count -eq 0)) { $items = Get-EaDefaultPovItem }
        Set-OptionControls $model.Options
        $script:C.txtRaw.Text = $Xml
        if ($null -eq $items) {
            # POV cannot be split into dimensions: fall back to raw editing.
            Set-PovRows @()
            Set-RawMode $true
            $script:S.LoadedXml = $Xml
        }
        else {
            Set-PovRows $items.ToArray()
            try { $script:S.LoadedXml = Get-EditorXml } catch { $script:S.LoadedXml = $Xml }
        }
        $script:C.tabEditor.IsEnabled = $true
        if ($Entry.IsDefault) { $info = 'Last-used EA settings for this user (read-only). Use "Save as new..." to turn them into a template.' }
        elseif ($Entry.IsNew) { $info = 'Fill in the POV and options, then Save.' }
        elseif ($Entry.Status -eq 'NotInIndex') { $info = "$($Entry.Key) exists but is missing from the SS_PREF index, so HFM does not list it. Save it, or use Tools > Repair template index." }
        else { $info = "$($Entry.Key)   |   template version $($model.Version)   |   $([Text.Encoding]::Unicode.GetByteCount($Xml)) bytes" }
        if ($null -eq $items) { $info += '   |   POV could not be split into dimensions; edit it in Raw XML.' }
        $script:C.txtTemplateInfo.Text = $info
        Update-Title
        Update-Validation
    }
    finally { $script:S.Loading = $false; Update-Buttons }
}

function Save-Current {
    # Returns $true when saved.
    $cur = $script:S.Current
    if (-not $cur -or $cur.IsDefault) { return $false }
    $errors = Get-EditorErrors
    if ($errors.Count) {
        Show-Message ("Cannot save '$($cur.Name)':`n`n" + ($errors -join "`n")) 'Validation' 'Warning'
        return $false
    }
    if ((Test-UserActive) -and -not (Confirm-Action "$($script:S.User) has an active HFM session. HFM may not pick up (or may overwrite) this change until they log in again.`n`nSave anyway?")) { return $false }
    $xml = Get-EditorXml
    $backup = Save-EaTemplate -App $script:S.App -User $script:S.User -Name $cur.Name -Xml $xml -New:([bool]$cur.IsNew)
    $script:S.Dirty = $false
    $key = Get-EaTemplateKey $cur.Name
    Update-TemplateList -SelectKey $key
    Set-Status "Saved $($cur.Name). Backup: $(Split-Path -Leaf $backup)"
    return $true
}

function Confirm-Discard {
    if (-not $script:S.Dirty -or -not $script:S.Current) { return $true }
    $cur = $script:S.Current
    if ($cur.IsDefault) {
        $r = [System.Windows.MessageBox]::Show($script:Window, 'Discard your edits to SS_DEFAULT? (Use "Save as new..." to keep them.)', 'Unsaved changes', 'OKCancel', 'Warning')
        if ($r -eq 'OK') { $script:S.Dirty = $false; return $true }
        return $false
    }
    $r = [System.Windows.MessageBox]::Show($script:Window, "Save changes to '$($cur.Name)'?", 'Unsaved changes', 'YesNoCancel', 'Warning')
    if ($r -eq 'Yes') { return (Save-Current) }
    if ($r -eq 'No') { $script:S.Dirty = $false; return $true }
    return $false
}

function Read-TemplateName([string]$Prompt, [string]$Initial, [string]$Title = 'Template name', [string]$AllowSameAs) {
    $d = New-XamlWindow 'NameDialog.xaml'
    $script:Dlg = $d
    $script:DlgAllow = $AllowSameAs
    $d.Window.Owner = $script:Window
    $d.Window.Title = $Title
    $d.C.txtPrompt.Text = $Prompt
    $d.C.txtName.Text = $Initial
    $d.C.btnOK.Add_Click({
            $name = $script:Dlg.C.txtName.Text.Trim()
            $err = Test-EaTemplateName $name
            if (-not $err -and $name -ne $script:DlgAllow) {
                $exists = @($script:C.lstTemplates.Items | Where-Object { -not $_.Tag.IsDefault -and $_.Tag.Name -eq $name }).Count -gt 0
                if ($exists) { $err = "A template named '$name' already exists for this user." }
            }
            if ($err) { $script:Dlg.C.txtError.Text = $err; return }
            $script:Dlg.Window.DialogResult = $true
        })
    $d.Window.Add_Loaded({ [void]$script:Dlg.C.txtName.Focus(); $script:Dlg.C.txtName.SelectAll() })
    if ($d.Window.ShowDialog()) { return $d.C.txtName.Text.Trim() }
    return $null
}

#endregion

#region Copy / import target dialog

function Show-TargetDialog([string]$Header, [string]$OkText, [string]$DefaultUser) {
    # Returns @{ App; Users; OnConflict } or $null.
    $d = New-XamlWindow 'CopyDialog.xaml'
    $script:Dlg = $d
    $d.Window.Owner = $script:Window
    $d.Window.Title = $OkText
    $d.C.txtHeader.Text = $Header
    $d.C.btnOK.Content = $OkText
    $script:DlgSelected = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if ($DefaultUser) { [void]$script:DlgSelected.Add($DefaultUser) }
    $script:DlgFillUsers = {
        $c = $script:Dlg.C
        foreach ($u in $c.lstUsers.SelectedItems) { [void]$script:DlgSelected.Add([string]$u) }
        $app = [string]$c.cmbTargetApp.SelectedItem
        $users = @(Get-HfmUser -App $app)
        foreach ($u in $script:DlgSelected) { if ($users -notcontains $u) { $users += $u } }
        $c.lstUsers.Items.Clear()
        foreach ($u in ($users | Sort-Object)) {
            [void]$c.lstUsers.Items.Add($u)
            if ($script:DlgSelected.Contains($u)) { [void]$c.lstUsers.SelectedItems.Add($u) }
        }
        $c.txtAppWarn.Visibility = if ($app -ne $script:S.App) { 'Visible' } else { 'Collapsed' }
    }
    $d.C.cmbTargetApp.ItemsSource = @(Get-HfmApp)
    $d.C.cmbTargetApp.SelectedItem = $script:S.App
    & $script:DlgFillUsers
    $d.C.cmbTargetApp.Add_SelectionChanged({ Invoke-Safe { & $script:DlgFillUsers } })
    $d.C.btnAddUser.Add_Click({
            $c = $script:Dlg.C
            $u = $c.txtNewUser.Text.Trim()
            if (-not $u) { return }
            if (-not $c.lstUsers.Items.Contains($u)) { [void]$c.lstUsers.Items.Add($u) }
            [void]$c.lstUsers.SelectedItems.Add($u)
            $c.txtNewUser.Text = ''
        })
    $d.C.btnOK.Add_Click({
            $c = $script:Dlg.C
            if ($c.lstUsers.SelectedItems.Count -eq 0) { $c.txtError.Text = 'Select at least one target user.'; return }
            $script:Dlg.Window.DialogResult = $true
        })
    if (-not $d.Window.ShowDialog()) { return $null }
    $c = $d.C
    $mode = if ($c.rbOverwrite.IsChecked) { 'Overwrite' } elseif ($c.rbRename.IsChecked) { 'Rename' } else { 'Skip' }
    return @{
        App        = [string]$c.cmbTargetApp.SelectedItem
        Users      = @($c.lstUsers.SelectedItems | ForEach-Object { [string]$_ })
        OnConflict = $mode
    }
}

function Show-AddResults($Outcome, [string]$Verb) {
    $lines = @($Outcome.Results | ForEach-Object { '{0,-14} {1} -> {2} / {3}' -f $_.Action, $_.Source, $_.User, $_.Name })
    $done = @($Outcome.Results | Where-Object Action -ne 'Skipped').Count
    $summary = "$Verb $done of $(@($Outcome.Results).Count) template copies."
    Set-Status "$summary Backup: $(Split-Path -Leaf $Outcome.Backup)"
    $text = $summary + "`n`n" + (($lines | Select-Object -First 40) -join "`n")
    if ($lines.Count -gt 40) { $text += "`n... ($($lines.Count - 40) more)" }
    Show-Message $text $Verb
}

#endregion

#region Actions

function Invoke-New {
    if (-not (Confirm-Discard)) { return }
    $name = Read-TemplateName -Prompt "Name of the new template for $($script:S.User):" -Initial '' -Title 'New template'
    if (-not $name) { return }
    $script:S.SuppressSelection = $true
    $script:C.lstTemplates.SelectedItems.Clear()
    $script:S.SuppressSelection = $false
    $xml = Get-NewEaTemplateXml
    $entry = [pscustomobject]@{ Name = $name; Key = (Get-EaTemplateKey $name); Status = 'OK'; IsDefault = $false; IsNew = $true }
    Show-Template $entry $xml
    $script:S.CurrentItem = $null
    $dsns = @($script:C.cmbDsn.ItemsSource)
    if ($dsns.Count) { $script:S.Loading = $true; $script:C.cmbDsn.Text = $dsns[0]; $script:S.Loading = $false }
    Update-Validation
    Update-Buttons
}

function Invoke-SaveAs {
    $cur = $script:S.Current
    if (-not $cur) { return }
    $errors = Get-EditorErrors
    if ($errors.Count) { Show-Message ("Fix these first:`n`n" + ($errors -join "`n")) 'Validation' 'Warning'; return }
    $initial = if ($cur.IsDefault) { '' } else { "$($cur.Name)_copy" }
    $name = Read-TemplateName -Prompt 'Save the current editor contents as a new template named:' -Initial $initial -Title 'Save as new template'
    if (-not $name) { return }
    $xml = Get-EditorXml
    $backup = Save-EaTemplate -App $script:S.App -User $script:S.User -Name $name -Xml $xml -New
    $script:S.Dirty = $false
    Update-TemplateList -SelectKey (Get-EaTemplateKey $name)
    Set-Status "Saved new template $name. Backup: $(Split-Path -Leaf $backup)"
}

function Invoke-Rename {
    $sel = @(Get-SelectedEntries)
    if ($sel.Count -ne 1 -or $sel[0].IsDefault) { return }
    if (-not (Confirm-Discard)) { return }
    $old = $sel[0].Name
    $new = Read-TemplateName -Prompt "Rename '$old' to:" -Initial $old -Title 'Rename template' -AllowSameAs $old
    if (-not $new -or $new -ceq $old) { return }
    $backup = Rename-EaTemplate -App $script:S.App -User $script:S.User -OldName $old -NewName $new
    $script:S.Dirty = $false
    Update-TemplateList -SelectKey (Get-EaTemplateKey $new)
    Set-Status "Renamed $old to $new. Backup: $(Split-Path -Leaf $backup)"
}

function Invoke-Delete {
    $sel = @(Get-SelectedEntries | Where-Object { -not $_.IsDefault })
    if (-not $sel.Count) { return }
    $names = ($sel | ForEach-Object Name) -join "`n  "
    if (-not (Confirm-Action "Delete $($sel.Count) template(s) for $($script:S.User)?`n`n  $names`n`nA backup is written first." 'Delete templates')) { return }
    $backup = Remove-EaTemplate -App $script:S.App -User $script:S.User -Keys @($sel | ForEach-Object Key)
    $script:S.Dirty = $false
    Update-TemplateList
    Set-Status "Deleted $($sel.Count) template(s). Backup: $(Split-Path -Leaf $backup)"
}

function Invoke-Copy {
    $sel = @(Get-SelectedEntries | Where-Object { $_.Status -ne 'MissingData' })
    if (-not $sel.Count) { return }
    if (-not (Confirm-Discard)) { return }
    $names = ($sel | ForEach-Object { if ($_.IsDefault) { 'SS_DEFAULT (as DEFAULT)' } else { $_.Name } }) -join ', '
    $t = Show-TargetDialog -Header "Copy from $($script:S.App) / $($script:S.User):`n$names" -OkText 'Copy'
    if (-not $t) { return }
    $outcome = Copy-EaTemplate -SourceApp $script:S.App -SourceUser $script:S.User -Keys @($sel | ForEach-Object Key) -TargetApp $t.App -TargetUsers $t.Users -OnConflict $t.OnConflict
    if ($t.App -eq $script:S.App -and $t.Users -contains $script:S.User) { Update-TemplateList -SelectKey $script:S.Current.Key }
    Show-AddResults $outcome 'Copied'
}

function Invoke-Export {
    $sel = @(Get-SelectedEntries | Where-Object { $_.Status -ne 'MissingData' })
    if (-not $sel.Count) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'EA template export (*.xml)|*.xml|All files (*.*)|*.*'
    $userPart = ($script:S.User -replace '[\\/:*?"<>|@ ]', '_')
    $dlg.FileName = if ($sel.Count -eq 1) { "$($script:S.App)_$($userPart)_$($sel[0].Name).xml" } else { "$($script:S.App)_$($userPart)_EATemplates.xml" }
    if (-not $dlg.ShowDialog($script:Window)) { return }
    Export-EaTemplate -App $script:S.App -User $script:S.User -Keys @($sel | ForEach-Object Key) -Path $dlg.FileName
    Set-Status "Exported $($sel.Count) template(s) to $($dlg.FileName)"
}

function Invoke-Import {
    if (-not (Confirm-Discard)) { return }
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'EA template export (*.xml)|*.xml|All files (*.*)|*.*'
    if (-not $dlg.ShowDialog($script:Window)) { return }
    $templates = Import-EaTemplateFile -Path $dlg.FileName
    if ($templates.Count -eq 0) { Show-Message 'The file contains no templates.'; return }
    $names = ($templates | ForEach-Object Name) -join ', '
    $t = Show-TargetDialog -Header "Import $($templates.Count) template(s) from $(Split-Path -Leaf $dlg.FileName):`n$names" -OkText 'Import' -DefaultUser $script:S.User
    if (-not $t) { return }
    $outcome = Add-EaTemplateSet -App $t.App -Users $t.Users -Templates $templates.ToArray() -OnConflict $t.OnConflict -Tag 'import'
    if ($t.App -eq $script:S.App -and $t.Users -contains $script:S.User) { Update-TemplateList }
    Show-AddResults $outcome 'Imported'
}

function Invoke-RepairIndex {
    if (-not (Confirm-Discard)) { return }
    if (-not (Confirm-Action "Rebuild the SS_PREF index for $($script:S.User)?`n`nTemplates missing from the index are added; index entries without data are removed.")) { return }
    Repair-EaIndex -App $script:S.App -User $script:S.User
    $script:S.Dirty = $false
    Update-TemplateList
}

function Invoke-Refresh {
    if (-not $script:S.App -or -not (Confirm-Discard)) { return }
    $script:S.Dirty = $false
    $script:S.ActiveUsers = @(Get-HfmActiveUser -App $script:S.App)
    $script:C.txtActiveWarn.Visibility = if (Test-UserActive) { 'Visible' } else { 'Collapsed' }
    $key = if ($script:S.Current -and -not $script:S.Current.IsNew) { $script:S.Current.Key } else { $null }
    Update-TemplateList -SelectKey $key
}

function Invoke-Restore {
    if (-not (Confirm-Discard)) { return }
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'EA Template Manager backup (*.csv)|*.csv'
    $dir = Join-Path $script:Root 'Backups'
    if (Test-Path $dir) { $dlg.InitialDirectory = $dir }
    if (-not $dlg.ShowDialog($script:Window)) { return }
    $groups = @(Import-Csv -Path $dlg.FileName | Group-Object App, Username, ParameterKey)
    if ($groups.Count -eq 0) { Show-Message 'The backup file is empty.'; return }
    $list = ($groups | Select-Object -First 15 | ForEach-Object { '  ' + ($_.Name -replace ', ', ' / ') }) -join "`n"
    if ($groups.Count -gt 15) { $list += "`n  ... ($($groups.Count - 15) more)" }
    $msg = "Put these $($groups.Count) key(s) back to the state saved in $(Split-Path -Leaf $dlg.FileName)?`n`n$list`n`nThe current rows are backed up first."
    if (-not (Confirm-Action $msg 'Restore backup')) { return }
    Restore-EaBackup -Path $dlg.FileName -Confirm:$false
    $script:S.Dirty = $false
    Set-Status "Restored $($groups.Count) key(s) from $(Split-Path -Leaf $dlg.FileName)."
    Invoke-Refresh
}

function Invoke-BackupAll {
    $path = Backup-EaRows -App $script:S.App -Tag 'full'
    Set-Status "Backed up all EA rows of $($script:S.App) to $path"
    Show-Message "Backed up all EA rows of $($script:S.App) to:`n$path"
}

#endregion

#region Main window

$main = New-XamlWindow 'MainWindow.xaml'
$script:Window = $main.Window
$script:C = $main.C
$script:Config = Get-Config

foreach ($p in $script:Config.ExtractFormats.PSObject.Properties) { [void]$script:C.cmbExtractFormat.Items.Add("$($p.Name) - $($p.Value)") }
foreach ($o in @($script:Config.LineItemOptions)) { [void]$script:C.cmbLineItem.Items.Add($o) }

# Editor change tracking
foreach ($n in $script:BoolOptions.Values) {
    $script:C[$n].Add_Checked({ Invoke-Safe { Set-Dirty } })
    $script:C[$n].Add_Unchecked({ Invoke-Safe { Set-Dirty } })
}
foreach ($n in $script:TextOptions.Values) {
    $ctl = $script:C[$n]
    if ($ctl -is [System.Windows.Controls.ComboBox]) {
        $ctl.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [System.Windows.Controls.TextChangedEventHandler] { Invoke-Safe { Set-Dirty } })
    }
    else { $ctl.Add_TextChanged({ Invoke-Safe { Set-Dirty } }) }
}
$script:C.txtRaw.Add_TextChanged({ if ($script:S.RawMode) { Invoke-Safe { Set-Dirty } } })

$script:C.btnAddCustom.Add_Click({
        Invoke-Safe {
            $nums = @($script:S.PovRows | Where-Object { $_.Dim -match '^C(\d+)$' } | ForEach-Object { [int]($_.Dim.Substring(1)) })
            $next = if ($nums.Count) { ($nums | Measure-Object -Maximum).Maximum + 1 } else { 1 }
            Add-PovRow "C$next" ''
            Set-Dirty
        }
    })
$script:C.btnRemoveCustom.Add_Click({
        Invoke-Safe {
            $items = @(Get-PovItems)
            $last = $items | Where-Object { $_.Dim -match '^C\d+$' } | Select-Object -Last 1
            if (-not $last) { return }
            Set-PovRows @($items | Where-Object { $_.Dim -ne $last.Dim })
            Set-Dirty
        }
    })

$script:C.tabEditor.Add_SelectionChanged({
        param($src, $e)
        if ($e.OriginalSource -ne $script:C.tabEditor) { return }
        if ($script:C.tabEditor.SelectedItem -eq $script:C.tabRaw -and -not $script:S.RawMode -and $script:S.BaseXml) {
            $script:S.Loading = $true
            try { $script:C.txtRaw.Text = Get-EditorXml } catch { $script:C.txtRaw.Text = "(cannot build XML: $($_.Exception.Message))" }
            finally { $script:S.Loading = $false }
        }
    })

$script:C.chkEditRaw.Add_Checked({
        if ($script:S.Loading) { return }
        try {
            $script:S.Loading = $true
            try { $script:C.txtRaw.Text = Get-EditorXml } finally { $script:S.Loading = $false }
            Set-RawMode $true
            Update-Validation
        }
        catch {
            Show-Message "The form cannot be converted to XML:`n`n$($_.Exception.Message)" 'Raw XML' 'Warning'
            Set-RawMode $false
        }
    })
$script:C.chkEditRaw.Add_Unchecked({
        if ($script:S.Loading) { return }
        try {
            $xml = $script:C.txtRaw.Text
            $model = ConvertFrom-EaTemplateXml $xml
            $items = Split-EaPov $model.Pov
            $script:S.Loading = $true
            try {
                $script:S.BaseXml = $xml
                Set-PovRows $items.ToArray()
                Set-OptionControls $model.Options
            }
            finally { $script:S.Loading = $false }
            Set-RawMode $false
            Set-Dirty
        }
        catch {
            Show-Message "The raw XML cannot be loaded into the form:`n`n$($_.Exception.Message)" 'Raw XML' 'Warning'
            Set-RawMode $true
        }
    })

# Selection
$script:C.cmbApp.Add_SelectionChanged({
        if ($script:S.SuppressApp) { return }
        Invoke-Safe {
            $app = [string]$script:C.cmbApp.SelectedItem
            if ($app -eq $script:S.App) { return }
            if (-not (Confirm-Discard)) {
                $script:S.SuppressApp = $true
                $script:C.cmbApp.SelectedItem = $script:S.App
                $script:S.SuppressApp = $false
                return
            }
            $script:S.Dirty = $false
            Select-App $app
        }
    })
$script:C.cmbUser.Add_SelectionChanged({
        Invoke-Safe { if ($script:C.cmbUser.SelectedItem) { Select-User ([string]$script:C.cmbUser.SelectedItem) } }
    })
$script:C.cmbUser.Add_PreviewKeyDown({
        param($src, $e)
        if ($e.Key -eq 'Return') { Invoke-Safe { Select-User $script:C.cmbUser.Text } }
    })
$script:C.btnLoadUser.Add_Click({
        Invoke-Safe {
            $u = $script:C.cmbUser.Text.Trim()
            if ($u -eq $script:S.User) {
                if (Confirm-Discard) { $script:S.Dirty = $false; Update-TemplateList -SelectKey $(if ($script:S.Current) { $script:S.Current.Key }) }
            }
            else { Select-User $u }
        }
    })

$script:C.lstTemplates.Add_SelectionChanged({
        if ($script:S.SuppressSelection) { return }
        Invoke-Safe {
            $item = $script:C.lstTemplates.SelectedItem
            if ($null -ne $item -and $item -eq $script:S.CurrentItem) { Update-Buttons; return }
            if ($null -eq $item -and $script:S.Current -and $script:S.Current.IsNew) { Update-Buttons; return }
            if (-not (Confirm-Discard)) {
                $script:S.SuppressSelection = $true
                $script:C.lstTemplates.SelectedItems.Clear()
                if ($script:S.CurrentItem) { $script:C.lstTemplates.SelectedItem = $script:S.CurrentItem }
                $script:S.SuppressSelection = $false
                return
            }
            if ($item) {
                Show-Template $item.Tag
                $script:S.CurrentItem = $item
            }
            else { Clear-Editor }
            Update-Buttons
        }
    })

# Buttons and menus
$script:C.btnSave.Add_Click({ Invoke-Safe { [void](Save-Current) } })
$script:C.btnRevert.Add_Click({
        Invoke-Safe {
            if ($script:S.Current -and -not $script:S.Current.IsNew) { $script:S.Dirty = $false; Show-Template $script:S.Current }
        }
    })
$script:C.btnSaveAs.Add_Click({ Invoke-Safe { Invoke-SaveAs } })
$script:C.btnNew.Add_Click({ Invoke-Safe { Invoke-New } })
$script:C.btnRename.Add_Click({ Invoke-Safe { Invoke-Rename } })
$script:C.btnDelete.Add_Click({ Invoke-Safe { Invoke-Delete } })
$script:C.btnCopy.Add_Click({ Invoke-Safe { Invoke-Copy } })
$script:C.btnExport.Add_Click({ Invoke-Safe { Invoke-Export } })
$script:C.btnImport.Add_Click({ Invoke-Safe { Invoke-Import } })
$script:C.mnuConnect.Add_Click({ Invoke-Safe { if (Confirm-Discard) { $script:S.Dirty = $false; Show-ConnectDialog } } })
$script:C.mnuBackupAll.Add_Click({ Invoke-Safe { Invoke-BackupAll } })
$script:C.mnuRestore.Add_Click({ Invoke-Safe { Invoke-Restore } })
$script:C.mnuOpenBackups.Add_Click({
        Invoke-Safe {
            $dir = Join-Path $script:Root 'Backups'
            if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
            Start-Process explorer.exe $dir
        }
    })
$script:C.mnuRepairIndex.Add_Click({ Invoke-Safe { Invoke-RepairIndex } })
$script:C.mnuRefresh.Add_Click({ Invoke-Safe { Invoke-Refresh } })
$script:C.mnuExit.Add_Click({ $script:Window.Close() })

$script:Window.Add_KeyDown({
        param($src, $e)
        if ($e.Key -eq 'F5') { Invoke-Safe { Invoke-Refresh } }
        elseif ($e.Key -eq 'S' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control' -and $script:C.btnSave.IsEnabled) { Invoke-Safe { [void](Save-Current) } }
    })
$script:Window.Add_Closing({
        param($src, $e)
        if ($script:S.Dirty -and -not (Confirm-Discard)) { $e.Cancel = $true }
    })
$script:Window.Add_ContentRendered({ Invoke-Safe { Show-ConnectDialog } })

Update-Buttons
[void]$script:Window.ShowDialog()
Disconnect-EaDatabase

#endregion
