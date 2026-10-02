param(
    [ValidateSet('Alle','Dokumentation','Krimi','Spielfilm','Serie')]
    [string]$Filter = 'Alle',

    [switch]$ExportOnly
)

$ErrorActionPreference = 'Stop'
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$IniPath = Join-Path $ScriptDir 'Internet TV (ger).ini'
$DatabasePath = Join-Path $env:APPDATA 'TV Movie\TV Movie ClickFinder\tvdaten.mdb'
$LogPath = Join-Path $ScriptDir 'ClickFinder_TV_Fehler.txt'


switch ($env:COMPUTERNAME) {
    'ERDMANN' { $DataDir = 'C:\Users\Erdmann\OneDrive\TV-Programm' }
    'LAPTOP'  { $DataDir = 'C:\Users\Erdmann\OneDrive\TV-Programm' }
    default   { throw "Unbekannter Rechnername: $env:COMPUTERNAME" }
}
$ExportPath = Join-Path $DataDir 'TV_Programm.json'
$GitHubRepoPath = 'C:\GitHub\tv-programm'
$GitHubJsonPath = Join-Path $GitHubRepoPath 'TV_Programm.json'
$GitHubPagesUrl = 'https://dieterwerdmann.github.io/tv-programm/TV_Programm.json'


function Remove-Diacritics([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object Text.StringBuilder
    foreach ($character in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($character)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($character)
        }
    }
    return $builder.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Normalize-Name([string]$Name) {
    $value = Remove-Diacritics $Name
    $value = $value.ToLowerInvariant()
    $value = $value -replace '\([^)]*\)', ''
    $value = $value -replace '\b(hd|uhd|sd|tv|fernsehen|fs|deutschland|de|live)\b', ''
    $value = $value -replace '[^a-z0-9]', ''
    return $value
}

function Read-IniChannels([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Senderliste fehlt: $Path"
    }

    $channels = @()
    $name = $null
    $url = $null

    foreach ($line in Get-Content -LiteralPath $Path -Encoding Default) {
        if ($line -match '^\[Channel\d+\]') {
            if ($name -and $url) {
                $channels += [pscustomobject]@{ Name=$name.Trim(); URL=$url.Trim() }
            }
            $name = $null
            $url = $null
        }
        elseif ($line -match '^Name=(.*)$') { $name = $matches[1] }
        elseif ($line -match '^URL=(.*)$')  { $url  = $matches[1] }
    }
    if ($name -and $url) {
        $channels += [pscustomobject]@{ Name=$name.Trim(); URL=$url.Trim() }
    }

    # Mehrfachfassungen wie (mul), (kla), (aud) nicht als eigene Hauptsender verwenden.
    return @($channels |
        Where-Object { $_.Name -notmatch '\((mul|kla|aud|deu|2)\)\s*$' } |
        Group-Object Name |
        ForEach-Object { $_.Group[0] })
}

function Get-ClickFinderData([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "ClickFinder-Datenbank wurde nicht gefunden:`n$Path`nBitte ClickFinder starten und die TV-Daten aktualisieren."
    }

    $connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$Path;Mode=Read;"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = @"
SELECT
    se.SenderKennung,
    se.Bezeichnung,
    sw.Beginn,
    sw.Ende,
    sw.Titel,
    sw.Kategorietext,
    sw.Genre,
    sw.Keywords,
    sw.KzFilm
FROM Sender AS se
INNER JOIN Sendungen AS sw
    ON se.SenderKennung = sw.SenderKennung
WHERE sw.Beginn <= ? AND sw.Ende > ?
ORDER BY se.Bezeichnung, sw.Beginn
"@
        $now = Get-Date
        [void]$command.Parameters.Add('pBeginn', [System.Data.OleDb.OleDbType]::Date)
        [void]$command.Parameters.Add('pEnde',   [System.Data.OleDb.OleDbType]::Date)
        $command.Parameters[0].Value = $now
        $command.Parameters[1].Value = $now

        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        # Das fuehrende Komma verhindert, dass PowerShell die DataTable
        # beim Rueckgabewert automatisch in einzelne DataRow-Objekte zerlegt.
        return ,$table
    }
    finally {
        $connection.Close()
        $connection.Dispose()
    }
}


function Get-ClickFinderDataRange([string]$Path, [datetime]$From, [datetime]$To) {
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "ClickFinder-Datenbank wurde nicht gefunden:`n$Path`nBitte ClickFinder starten und die TV-Daten aktualisieren."
    }

    $connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$Path;Mode=Read;"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = @"
SELECT
    se.SenderKennung,
    se.Bezeichnung,
    sw.Beginn,
    sw.Ende,
    sw.Titel,
    sw.Kategorietext,
    sw.Genre,
    sw.Keywords,
    sw.KzFilm
FROM Sender AS se
INNER JOIN Sendungen AS sw
    ON se.SenderKennung = sw.SenderKennung
WHERE sw.Beginn < ? AND sw.Ende > ?
ORDER BY sw.Beginn, se.Bezeichnung
"@
        [void]$command.Parameters.Add('pTo',   [System.Data.OleDb.OleDbType]::Date)
        [void]$command.Parameters.Add('pFrom', [System.Data.OleDb.OleDbType]::Date)
        $command.Parameters[0].Value = $To
        $command.Parameters[1].Value = $From

        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
        $table = New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        return ,$table
    }
    finally {
        $connection.Close()
        $connection.Dispose()
    }
}

function Export-TVProgramJson([string]$DatabasePath, $IniChannels, [string]$Path) {
    $from = (Get-Date).Date
    $to = $from.AddDays(2)

    $table = Get-ClickFinderDataRange $DatabasePath $from $to
    $mapping = Build-Mapping $IniChannels $table

    # Die Privatsender werden fuer Android/JSON direkt ueber die
    # ClickFinder-Senderkennung auf die bestaetigten Anzeigenamen abgebildet.
    # Dadurch sind keine kuenstlichen DVBViewer-/INI-Streams erforderlich.
    $privateSender = @{
        'XXP'             = 'DMAX HD'
        'EUROSPORT'       = 'Eurosport 1 HD'
        'HGTV'            = 'HGTV HD'
        'KABEL'           = 'Kabel Eins Deutschland HD'
        'NICK'            = 'Nick/CC+1 HD'
        'RTL NITRO'       = 'NITRO HD'
        'PRO7'            = 'Pro7 Deutschland HD'
        'PROSIEBEN MAXX'  = 'ProSieben MAXX HD'
        'RTL'             = 'RTL Deutschland HD'
        'RTL PLUS'        = 'RTLup HD'
        'RTL II'          = 'RTLZWEI Deutschland HD'
        'SAT1'            = 'Sat.1 Deutschland HD'
        'SAT.1 GOLD'      = 'SAT.1 Gold HD'
        'SIXX'            = 'Sixx HD'
        'S RTL'           = 'Super RTL Deutschland HD'
        'TELE5'           = 'TELE 5 HD'
        'VOX'             = 'Vox Deutschland HD'
        'VOXUP'           = 'VOXup HD'
        'N24'             = 'WELT HD'
    }

    $programs = @()

    # Normale INI-/DVBViewer-Sender.
    # Privatsender werden hier uebersprungen und danach mit ihren
    # bestaetigten Smartphone-Anzeigenamen direkt aus ClickFinder erzeugt.
    foreach ($entry in @($mapping.Entries)) {
        $senderKey = ([string]$entry.SenderKennung).Trim().ToUpperInvariant()

        if ($privateSender.ContainsKey($senderKey)) {
            continue
        }

        $kind = Get-Category $entry.Row
        if (-not $kind) { continue }

        $programs += [pscustomobject]@{
            Art     = $kind
            Sender  = [string]$entry.IniName
            Beginn  = ([datetime]$entry.Row.Beginn).ToString('yyyy-MM-ddTHH:mm:ss')
            Ende    = ([datetime]$entry.Row.Ende).ToString('yyyy-MM-ddTHH:mm:ss')
            Sendung = [string]$entry.Row.Titel
        }
    }

    # Privatsender direkt aus den ClickFinder-Daten.
    $privateRows = if ($table -is [System.Data.DataTable]) {
        @($table.Rows)
    }
    else {
        @($table)
    }

    foreach ($row in $privateRows) {
        $senderKey = ([string]$row.SenderKennung).Trim().ToUpperInvariant()

        if (-not $privateSender.ContainsKey($senderKey)) {
            continue
        }

        $kind = Get-Category $row
        if (-not $kind) { continue }

        $programs += [pscustomobject]@{
            Art     = $kind
            Sender  = [string]$privateSender[$senderKey]
            Beginn  = ([datetime]$row.Beginn).ToString('yyyy-MM-ddTHH:mm:ss')
            Ende    = ([datetime]$row.Ende).ToString('yyyy-MM-ddTHH:mm:ss')
            Sendung = [string]$row.Titel
        }
    }

    $order = @{ Dokumentation=1; Krimi=2; Spielfilm=3; Serie=4 }
    $programs = @($programs | Sort-Object Beginn, @{Expression={$order[$_.Art]}}, Sender)

    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) {
        [void](New-Item -ItemType Directory -Path $folder -Force)
    }

    $payload = [pscustomobject]@{
        Aktualisiert = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
        ZeitraumVon  = $from.ToString('yyyy-MM-ddTHH:mm:ss')
        ZeitraumBis  = $to.ToString('yyyy-MM-ddTHH:mm:ss')
        Anzahl       = $programs.Count
        Sendungen    = $programs
    }

    $payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
    return $programs.Count
}


function Prepare-GitHubRepo([string]$RepoPath) {
    if (-not (Test-Path -LiteralPath $RepoPath)) {
        throw "GitHub-Arbeitskopie wurde nicht gefunden: $RepoPath"
    }

    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if (-not $git) {
        throw "git.exe wurde nicht gefunden. Bitte Git fuer Windows installieren."
    }

    & $git.Source -C $RepoPath pull --rebase --autostash --quiet
    if ($LASTEXITCODE -ne 0) {
        throw "git pull ist fehlgeschlagen."
    }
}

function Publish-GitHubPages([string]$RepoPath, [string]$PagesUrl) {
    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if (-not $git) {
        return [pscustomobject]@{
            Ok      = $false
            Changed = $false
            Message = 'git.exe wurde nicht gefunden.'
        }
    }

    $statusText = (& $git.Source -C $RepoPath status --porcelain -- 'TV_Programm.json' 2>$null) -join "`n"

    if ([string]::IsNullOrWhiteSpace($statusText)) {
        return [pscustomobject]@{
            Ok      = $true
            Changed = $false
            Message = 'bereits aktuell'
        }
    }

    & $git.Source -C $RepoPath add -- 'TV_Programm.json'
    if ($LASTEXITCODE -ne 0) {
        return [pscustomobject]@{
            Ok      = $false
            Changed = $true
            Message = 'git add ist fehlgeschlagen.'
        }
    }

    & $git.Source -C $RepoPath commit -m 'TV-Programm aktualisiert' --quiet
    if ($LASTEXITCODE -ne 0) {
        return [pscustomobject]@{
            Ok      = $false
            Changed = $true
            Message = 'git commit ist fehlgeschlagen.'
        }
    }

    & $git.Source -C $RepoPath push --quiet
    if ($LASTEXITCODE -ne 0) {
        return [pscustomobject]@{
            Ok      = $false
            Changed = $true
            Message = 'git push ist fehlgeschlagen.'
        }
    }

    return [pscustomobject]@{
        Ok      = $true
        Changed = $true
        Message = ('aktualisiert -> {0}' -f $PagesUrl)
    }
}

function Get-Category($row) {
    $title = [string]$row.Titel
    $genre = [string]$row.Genre
    $category = [string]$row.Kategorietext
    $keywords = [string]$row.Keywords
    $all = ($title + ' ' + $genre + ' ' + $category + ' ' + $keywords).ToLowerInvariant()
    $isFilm = $false
    if ($row.KzFilm -ne [DBNull]::Value) { $isFilm = [bool]$row.KzFilm }

    # Krimi vor Spielfilm, weil ein Kriminalfilm häufig zugleich KzFilm=True ist.
    if ($all -match 'krimi|kriminal|crime|detective|polizei|thriller|politthriller|mystery') {
        return 'Krimi'
    }

    if ($isFilm -or $all -match 'spielfilm|fernsehfilm|kinofilm|drama|komödie|western|actionfilm|liebesfilm|romanze|abenteuerfilm') {
        return 'Spielfilm'
    }

    if ($all -match 'serie|serien|soap|sitcom|telenovela|daily|episode|staffel|folge') {
        return 'Serie'
    }

    if ($all -match 'dokumentation|dokumentar|reportage|natur|tierwelt|tiere|geschichte|wissenschaft|wissen|reise|kultur|gesellschaft und soziales|biografie|portrait|porträt|technik|umwelt|archäologie') {
        return 'Dokumentation'
    }

    if (([string]$row.SenderKennung).Trim().ToUpperInvariant() -eq 'NTV') {
        return 'Dokumentation'
    }

    return $null
}

function Get-DVBViewerPath {
    $process = Get-Process -Name 'DVBViewer' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($process -and $process.Path -and (Test-Path -LiteralPath $process.Path)) {
        return $process.Path
    }

    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'DVBViewer\DVBViewer.exe'),
        (Join-Path $env:ProgramFiles 'DVBViewer\DVBViewer.exe'),
        'C:\Program Files (x86)\DVBViewer\DVBViewer.exe',
        'C:\Program Files\DVBViewer\DVBViewer.exe'
    ) | Where-Object { $_ }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

function Get-EdgeTopLevelFenster {

    if (-not ('TVProgrammEdgeWindows' -as [type])) {

        Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class TVProgrammEdgeWindows
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(
        EnumWindowsProc lpEnumFunc,
        IntPtr lParam
    );

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(
        IntPtr hWnd
    );

    [DllImport("user32.dll")]
    public static extern bool IsWindow(
        IntPtr hWnd
    );

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(
        IntPtr hWnd
    );

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(
        IntPtr hWnd,
        StringBuilder lpString,
        int nMaxCount
    );

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint lpdwProcessId
    );

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool PostMessage(
        IntPtr hWnd,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam
    );
}
"@
    }

    $Liste = New-Object System.Collections.Generic.List[object]

    $Callback = [TVProgrammEdgeWindows+EnumWindowsProc]{

        param($hWnd, $lParam)

        if ([TVProgrammEdgeWindows]::IsWindowVisible($hWnd)) {

            $Laenge = [TVProgrammEdgeWindows]::GetWindowTextLength($hWnd)

            if ($Laenge -gt 0) {

                $Text = New-Object System.Text.StringBuilder ($Laenge + 1)

                [void][TVProgrammEdgeWindows]::GetWindowText(
                    $hWnd,
                    $Text,
                    $Text.Capacity
                )

                [uint32]$ProcessId = 0

                [void][TVProgrammEdgeWindows]::GetWindowThreadProcessId(
                    $hWnd,
                    [ref]$ProcessId
                )

                $Prozess = Get-Process `
                    -Id $ProcessId `
                    -ErrorAction SilentlyContinue

                if ($Prozess -and $Prozess.ProcessName -eq 'msedge') {

                    $Liste.Add(
                        [pscustomobject]@{
                            HWND      = $hWnd.ToInt64()
                            ProcessId = $ProcessId
                            Titel     = $Text.ToString()
                        }
                    )
                }
            }
        }

        return $true
    }

    [void][TVProgrammEdgeWindows]::EnumWindows(
        $Callback,
        [IntPtr]::Zero
    )

    return $Liste.ToArray()
}


function Stop-ZattooFenster {

    if (-not $script:ZattooWindowHandle) {
        return
    }

    $Handle = [IntPtr]$script:ZattooWindowHandle

    if (-not [TVProgrammEdgeWindows]::IsWindow($Handle)) {
        $script:ZattooWindowHandle = $null
        return
    }

    Write-Host (
        "Vorhandenes Zattoo-Fenster wird geschlossen: HWND {0}" -f
        $script:ZattooWindowHandle
    ) -ForegroundColor DarkYellow

    $WM_CLOSE = 0x0010

    $Gesendet = [TVProgrammEdgeWindows]::PostMessage(
        $Handle,
        $WM_CLOSE,
        [IntPtr]::Zero,
        [IntPtr]::Zero
    )

    if (-not $Gesendet) {
        throw "WM_CLOSE konnte nicht an das Zattoo-Fenster gesendet werden."
    }

    $Ende = (Get-Date).AddSeconds(5)

    while (
        [TVProgrammEdgeWindows]::IsWindow($Handle) -and
        (Get-Date) -lt $Ende
    ) {
        Start-Sleep -Milliseconds 200
    }

    if ([TVProgrammEdgeWindows]::IsWindow($Handle)) {
        throw "Das gemerkte Zattoo-Fenster konnte nicht sauber geschlossen werden."
    }

    $script:ZattooWindowHandle = $null
}



function Start-DVBViewerChannel([string]$ChannelName) {
    $exe = Get-DVBViewerPath

    if (-not $exe) {
        Write-Host 'DVBViewer.exe wurde nicht gefunden.' -ForegroundColor Red
        return
    }

    # Wechsel Zattoo -> DVBViewer:
    # Nur ein erkanntes Zattoo-Fenster schließen.
    # Andere Edge-Fenster bleiben unangetastet.
    Stop-ZattooFenster

    $argument = '-c' + $ChannelName

    Write-Host ("Umschalten auf {0} ..." -f $ChannelName) -ForegroundColor Green

    # Ist DVBViewer bereits offen, übernimmt die vorhandene Instanz
    # den neuen Kanal. DVBViewer wird deshalb hier NICHT beendet.
    Start-Process -FilePath $exe -ArgumentList ('"{0}"' -f $argument)
}

function Find-IniChannelForClickFinder($row, $iniChannels) {
    $key = ([string]$row.SenderKennung).Trim().ToUpperInvariant()
    $label = ([string]$row.Bezeichnung).Trim()

    # Bevorzugte DVBViewer-Namen. Der erste vorhandene Eintrag wird verwendet.
    $preferred = @{
        'ARD'=@('Das Erste HD','Das Erste','ARD')
        'ZDF'=@('ZDF HD','ZDF')
        'ARTE'=@('arte HD','Arte HD','arte','Arte')
        '3SAT'=@('3sat HD','3sat')
        'ZDFNEO'=@('zdf_neo HD','ZDFneo HD','zdf_neo','ZDFneo')
        'ZDFINFO'=@('ZDFinfo HD','ZDFinfo')
        'ONE'=@('ONE HD','ONE')
        'PHOENIX'=@('phoenix','Phoenix')
        'ARDALPHA'=@('ARD-alpha HD','ARD-alpha','ard alpha')
        'KIKA'=@('KiKa HD','KiKa','KIKA')
        'WDR'=@('WDR Fernsehen','WDR')
        'NDR'=@('NDR FS HH HD','NDR FS NDS HD','NDR FS SH HD','NDR FS MV HD','NDR Fernsehen')
        'HR'=@('hr-fernsehen HD','hr-fernsehen','HR Fernsehen')
        'BR'=@('BR Fernsehen Nord HD','BR Fernsehen Süd HD','BR Fernsehen')
        'MDR'=@('MDR Sachsen HD','MDR S-Anhalt HD','MDR Thüringen HD','MDR Fernsehen')
        'SWR'=@('SWR BW HD','SWR RP HD','SWR Fernsehen')
        'RBB'=@('rbb Berlin HD','rbb Brandenburg HD','rbb Fernsehen')
        'RTL'=@('RTL','RTL HD')
        'SAT1'=@('SAT.1','SAT1','Sat.1')
        'PRO7'=@('ProSieben','PRO7','Pro7')
        'RTL2'=@('RTL II','RTL2','RTLZWEI')
        'KABEL1'=@('Kabel Eins','kabel eins','Kabel 1')
        'VOX'=@('VOX','Vox')
        'COMEDYCENTRAL'=@('Comedy Central')
        'MTV'=@('MTV')
        'WELT'=@('Welt HD','Welt')
        'N24DOKU'=@('N24 Doku HD','N24 Doku')
        'BIBELTV'=@('Bibel TV')
        'DMAX'=@('DMAX')
        'SIXX'=@('Sixx','SIXX')
        'TELE5'=@('Tele 5','TELE 5')
        'SUPER RTL'=@('Super RTL','SUPER RTL')
        'RTLPLUS'=@('RTLplus','RTLup')
        'NICKELODEON'=@('Nickelodeon')
        'DISNEYCHANNEL'=@('Disney Channel')
        'KABEL1DOKU'=@('Kabel Eins Doku','kabel eins Doku','Kabel1 Doku','Kabel1 Doku HD')
        'NITRO'=@('NITRO','RTL Nitro')
        'SAT1GOLD'=@('SAT.1 Gold','Sat.1 Gold')
        'PRO7MAXX'=@('ProSieben MAXX','Pro7 MAXX')
        'DF1'=@('DF1 HD','DF1')
        'ANIXE'=@('Anixe Serie','ANIXE HD','ANIXE')
    }

    if ($preferred.ContainsKey($key)) {
        foreach ($wanted in $preferred[$key]) {
            $found = $iniChannels | Where-Object { $_.Name -ieq $wanted } | Select-Object -First 1
            if ($found) { return $found }
        }
    }

    # Danach exakte, normalisierte Namenssuche.
    $targets = @((Normalize-Name $label), (Normalize-Name $key)) | Where-Object { $_ }
    foreach ($target in $targets) {
        $found = $iniChannels | Where-Object { (Normalize-Name $_.Name) -eq $target } | Select-Object -First 1
        if ($found) { return $found }
    }

    # Keine Teilnamensuche mehr. Sie kann verschiedene Sender verwechseln,
    # z.B. "Kabel Eins" mit "Kabel Eins Doku".
    # Nicht eindeutig zuordenbare Sender werden lieber weggelassen.
    return $null
}

function Build-Mapping($iniChannels, $dbRows) {
    # Senderzuordnung cachen: Ein Sender wird nur einmal gegen die INI-Liste
    # aufgeloest. Das beschleunigt besonders den 2-Tage-Android-Export deutlich.
    $result = [System.Collections.Generic.List[object]]::new()
    $unmapped = [System.Collections.Generic.HashSet[string]]::new()
    $channelCache = @{}

    # DataTable explizit in einzelne DataRow-Objekte umwandeln.
    if ($dbRows -is [System.Data.DataTable]) {
        $rows = @($dbRows.Rows)
    }
    else {
        $rows = @($dbRows)
    }

    foreach ($row in $rows) {
        $cacheKey = ('{0}|{1}' -f ([string]$row.SenderKennung).Trim().ToUpperInvariant(),
                                  ([string]$row.Bezeichnung).Trim())

        if ($channelCache.ContainsKey($cacheKey)) {
            $channel = $channelCache[$cacheKey]
        }
        else {
            $channel = Find-IniChannelForClickFinder $row $iniChannels
            $channelCache[$cacheKey] = $channel
        }

        if ($null -ne $channel) {
            $result.Add([pscustomobject]@{
                IniName = [string]$channel.Name
                URL = [string]$channel.URL
                SenderKennung = [string]$row.SenderKennung
                Row = $row
            })
        }
        else {
            if (([string]$row.SenderKennung).Trim().ToUpperInvariant() -eq 'NTV') {
                $result.Add([pscustomobject]@{
                    IniName        = 'NTV'
                    URL            = ''
                    SenderKennung  = [string]$row.SenderKennung
                    Row            = $row
                })
            }
            else {
                [void]$unmapped.Add(('{0} [{1}]' -f [string]$row.Bezeichnung, [string]$row.SenderKennung))
            }
        }
    }

    [pscustomobject]@{
        Entries = @($result)
        Unmapped = @($unmapped | Sort-Object)
    }
}

# Senderkennungen zentral verwalten.
# NTV bleibt separat in der Hauptliste; die 21 Sender hier bilden das Untermenue "Privatsender".
$script:ZattooKanaele = [ordered]@{
    'DMAX HD'                  = 'dmax'
    'Eurosport 1 HD'           = 'eurosport1'
    'HGTV HD'                  = 'hgtv_de'
    'Kabel Eins Deutschland HD'= 'kabel_eins_deutschland'
    'Nick/CC+1 HD'             = 'nick_cc_plus1'
    'NITRO HD'                 = 'rtlnitro_de'
    'Pro7 Deutschland HD'      = 'pro7_deutschland'
    'ProSieben MAXX HD'        = 'pro7maxx'
    'RTL Deutschland HD'       = 'rtl_deutschland'
    'RTLup HD'                 = 'rtl_up_de'
    'RTLZWEI Deutschland HD'   = 'rtl2_deutschland'
    'Sat.1 Deutschland HD'     = 'sat1_deutschland'
    'SAT.1 Gold HD'            = 'sat1gold'
    'Sixx HD'                  = 'sixx_deutschland'
    'Super RTL Deutschland HD' = 'super_rtl_deutschland'
    'TELE 5 HD'                = 'tele-5'
    'TOGGO plus HD'            = 'toggo_plus'
    'Vox Deutschland HD'       = 'vox_deutschland'
    'VOXup HD'                 = 'vox_up_de'
    'WELT HD'                  = 'welt'
    'XPLORE HD'                = 'xplore_de'
}

# Alle 21 Sender des Untermenues werden ueber dieselbe Start-/Fensterfunktion angesprochen.
$script:AktiveZattooSender = @($script:ZattooKanaele.Keys)

function Set-ZattooFenster {
    Add-Type -AssemblyName System.Windows.Forms

    if (-not ('ZattooFenster' -as [type])) {
        Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class ZattooFenster
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(
        IntPtr hWnd,
        int nCmdShow
    );

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags
    );
}
'@
    }

    $Treffer = New-Object System.Collections.Generic.List[object]

    $Callback = [ZattooFenster+EnumWindowsProc]{
        param($hWnd, $lParam)

        if ([ZattooFenster]::IsWindowVisible($hWnd)) {
            $Text = New-Object System.Text.StringBuilder 512
            [void][ZattooFenster]::GetWindowText($hWnd, $Text, $Text.Capacity)

            if ($Text.ToString() -like 'Sender*Microsoft*Edge*') {
                $Treffer.Add([pscustomobject]@{ Handle = $hWnd })
            }
        }

        return $true
    }

    [void][ZattooFenster]::EnumWindows($Callback, [IntPtr]::Zero)

    $Fenster = $Treffer | Select-Object -First 1

    if ($Fenster) {
        $Screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

        $Links  = [int]($Screen.Width  * 0.731)
        $Oben   = $Screen.Top
        $Breite = [int]($Screen.Width  * 0.269)
        $Hoehe  = [int]($Screen.Height * 0.522)

        [void][ZattooFenster]::ShowWindow(
            [IntPtr]$Fenster.Handle,
            9
        )

        Start-Sleep -Milliseconds 500

        [void][ZattooFenster]::SetWindowPos(
            [IntPtr]$Fenster.Handle,
            [IntPtr]::Zero,
            $Links,
            $Oben,
            $Breite,
            $Hoehe,
            0x0040
        )
    }
}




function Stop-DVBViewerFuerZattoo {
    $DVBProzesse = @(
        Get-Process -Name 'DVBViewer' -ErrorAction SilentlyContinue
    )

    if ($DVBProzesse.Count -eq 0) {
        return
    }

    Write-Host 'DVBViewer wird vor dem Start von Zattoo geschlossen ...' -ForegroundColor DarkYellow

    foreach ($Prozess in $DVBProzesse) {
        if ($Prozess.MainWindowHandle -ne 0) {
            [void]$Prozess.CloseMainWindow()
        }
    }

    $Ende = (Get-Date).AddSeconds(5)

    do {
        Start-Sleep -Milliseconds 200
        $NochOffen = @(Get-Process -Name 'DVBViewer' -ErrorAction SilentlyContinue)
    } while ($NochOffen.Count -gt 0 -and (Get-Date) -lt $Ende)

    if ($NochOffen.Count -gt 0) {
        Write-Host 'DVBViewer reagiert nicht auf normales Schließen – Prozess wird beendet.' -ForegroundColor DarkYellow
        $NochOffen | Stop-Process -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 500
    }
}


function Start-ZattooSender([string]$SenderName) {

    if (-not $script:ZattooKanaele.Contains($SenderName)) {
        throw "Keine verifizierte Zattoo-Kennung fuer: $SenderName"
    }

    if ($SenderName -notin $script:AktiveZattooSender) {
        throw "Sender noch nicht fuer Test freigegeben: $SenderName"
    }

    $Edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

    if (-not (Test-Path -LiteralPath $Edge)) {
        throw "Microsoft Edge wurde nicht gefunden: $Edge"
    }

    $SenderKennung = $script:ZattooKanaele[$SenderName]
    $PrivatUrl = 'https://zattoo.com/channels/favorites?channel=' + $SenderKennung

    # DVBViewer -> Zattoo
    Stop-DVBViewerFuerZattoo

    # Zattoo -> Zattoo
    Stop-ZattooFenster

    # Nur bereits vorhandene sichtbare Senderfenster merken.
    $SenderFensterVorher = @(
        Get-EdgeTopLevelFenster |
            Where-Object {
                $_.Titel -like 'Sender -*Microsoft*Edge'
            } |
            Select-Object -ExpandProperty HWND
    )

    Write-Host (
        "{0} wird ueber Zattoo geoeffnet ..." -f $SenderName
    ) -ForegroundColor Green

    Start-Process `
        -FilePath $Edge `
        -ArgumentList @(
            '--new-window',
            $PrivatUrl
        )

    # Nur prüfen, ob überhaupt ein neues Senderfenster entstanden ist.
    $Ende = (Get-Date).AddSeconds(10)
    $StartFensterGefunden = $false

    do {
        Start-Sleep -Milliseconds 250

        $NeueSenderFenster = @(
            Get-EdgeTopLevelFenster |
                Where-Object {
                    $_.Titel -like 'Sender -*Microsoft*Edge' -and
                    $_.HWND -notin $SenderFensterVorher
                }
        )

        if ($NeueSenderFenster.Count -ge 1) {
            $StartFensterGefunden = $true
        }

    } while (
        -not $StartFensterGefunden -and
        (Get-Date) -lt $Ende
    )

    if (-not $StartFensterGefunden) {
        throw "Nach dem Zattoo-Start wurde kein neues Senderfenster erkannt."
    }

    # Edge darf jetzt vollständig laden und ggf. sein HWND wechseln.
    Start-Sleep -Seconds 5

    Set-ZattooFenster

    # Nach Laden und Positionierung das JETZT tatsächlich sichtbare
    # neue Senderfenster erneut bestimmen.
    Start-Sleep -Milliseconds 500

    $EndgueltigeSenderFenster = @(
        Get-EdgeTopLevelFenster |
            Where-Object {
                $_.Titel -like 'Sender -*Microsoft*Edge' -and
                $_.HWND -notin $SenderFensterVorher
            }
    )

    if ($EndgueltigeSenderFenster.Count -ne 1) {

        Write-Host "`nAktuelle neue Senderfenster:" -ForegroundColor Yellow

        $EndgueltigeSenderFenster |
            Format-Table HWND, ProcessId, Titel -AutoSize

        throw (
            "Endgueltiges Zattoo-Fenster nicht eindeutig: {0} Treffer." -f
            $EndgueltigeSenderFenster.Count
        )
    }

    $script:ZattooWindowHandle = [int64]$EndgueltigeSenderFenster[0].HWND

    Write-Host (
        "Zattoo-Fenster endgueltig gespeichert: HWND {0} | {1}" -f
        $script:ZattooWindowHandle,
        $EndgueltigeSenderFenster[0].Titel
    ) -ForegroundColor DarkGreen
}

try {
    $scriptStart = Get-Date
    Remove-Item -LiteralPath $LogPath -Force -ErrorAction SilentlyContinue

    $iniChannels = Read-IniChannels $IniPath

    # Zusaetzlicher Export fuer die Android-Anzeige:
    # komplettes Programm fuer heute und morgen als Nutzdaten auf E:.
    $exportCount = Export-TVProgramJson $DatabasePath $iniChannels $ExportPath

    # Fuer GitHub Pages nur eine Kopie der Nutzdaten in die lokale GitHub-Arbeitskopie schreiben.
    if (-not (Test-Path -LiteralPath $GitHubRepoPath)) {
        throw "GitHub-Arbeitskopie wurde nicht gefunden: $GitHubRepoPath"
    }
    Copy-Item -LiteralPath $ExportPath -Destination $GitHubJsonPath -Force

    # Unbeaufsichtigter Modus fuer die automatische Android-/GitHub-Aktualisierung.
    if ($ExportOnly) {

        $syncScript = Join-Path $ScriptDir 'GitHub_TV_Sync.ps1' 
        $pwsh = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source

        if (-not $pwsh) {
            throw 'PowerShell 7 (pwsh.exe) wurde fuer den GitHub-Sync nicht gefunden.'
        }

        if (-not (Test-Path -LiteralPath $syncScript)) {
            throw "GitHub-Sync-Skript wurde nicht gefunden: $syncScript"
        }

        Write-Host ("Android-Export: {0} Sendungen -> {1}" -f $exportCount, $ExportPath)
        Write-Host 'GitHub-Synchronisation wird ausgefuehrt ...'

        & $pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File $syncScript
        $syncExitCode = $LASTEXITCODE

        if ($syncExitCode -ne 0) {
            throw "GitHub-Synchronisation fehlgeschlagen (ExitCode $syncExitCode)."
        }

        Write-Host 'OK - Android-Export und GitHub-Synchronisation abgeschlossen.' -ForegroundColor Green
        exit 0
    }

    $dbTable = Get-ClickFinderData $DatabasePath
    if ($dbTable -isnot [System.Data.DataTable]) {
        throw "Interner Datenbankfehler: Erwartet wurde eine DataTable, erhalten: $($dbTable.GetType().FullName)"
    }
    $mapping = Build-Mapping $iniChannels $dbTable
    $mapped = @($mapping.Entries)

    $items = @()
    foreach ($entry in $mapped) {
        $kind = Get-Category $entry.Row
        if (-not $kind) { continue }
        if ($Filter -ne 'Alle' -and $kind -ne $Filter) { continue }

        $items += [pscustomobject]@{
            Art = $kind
            Sender = [string]$entry.IniName
            Beginn = [datetime]$entry.Row.Beginn
            Ende = [datetime]$entry.Row.Ende
            Sendung = [string]$entry.Row.Titel
        }
    }

    $order = @{ Dokumentation=1; Krimi=2; Spielfilm=3; Serie=4 }
    $items = @($items | Sort-Object @{Expression={$order[$_.Art]}}, Sender, Beginn)

    Clear-Host
    Write-Host ("AKTUELLES TV-PROGRAMM  {0:dd.MM.yyyy HH:mm}" -f (Get-Date))
    Write-Host ("Quelle: TV Movie ClickFinder | Filter: {0}" -f $Filter)
    Write-Host ("Android-Export: {0} Sendungen -> {1}" -f $exportCount, $ExportPath)
    Write-Host ("GitHub Pages:  Aktualisierung laeuft im Hintergrund -> {0}" -f $GitHubPagesUrl) -ForegroundColor DarkGray
    Write-Host ("Laufzeit bis zur Anzeige: {0:N1} Sekunden" -f (((Get-Date) - $scriptStart).TotalSeconds)) -ForegroundColor DarkGray

    # GitHub-Synchronisation im Hintergrund ueber ein separates Skript.
    $syncScript = Join-Path $ScriptDir 'GitHub_TV_Sync.ps1'
    $pwsh = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source

    if ($pwsh -and (Test-Path -LiteralPath $syncScript)) {
        Start-Process -FilePath $pwsh `
            -ArgumentList @('-NoLogo','-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $syncScript)) `
            -WindowStyle Hidden
    }
    else {
        Write-Host 'Hinweis: GitHub-Hintergrund-Sync konnte nicht gestartet werden.' -ForegroundColor Yellow
    }

    Write-Host ('-' * 105)
    Write-Host ''


    if ($items.Count -eq 0) {
        Write-Host 'Zurzeit wurden keine passenden Sendungen gefunden.' -ForegroundColor Yellow
        Write-Host ''
        Write-Host ("ClickFinder lieferte {0} aktuell laufende Programme; {1} Programme wurden Ihrer INI-Liste zugeordnet." -f $dbTable.Rows.Count, $mapped.Count)
        Write-Host ("Sender in Ihrer INI-Datei: {0}" -f $iniChannels.Count)
        Write-Host ("Datentyp der Datenbankabfrage: {0}" -f $dbTable.GetType().FullName)
        if ($mapping.Unmapped.Count -gt 0) {
            Write-Host ""
            Write-Host "Nicht zugeordnete ClickFinder-Sender:" -ForegroundColor Yellow
            $mapping.Unmapped | ForEach-Object { Write-Host ("  " + $_) }
        }
        Read-Host 'ENTER beendet'
        exit 0
    }

    $numbered = @()
    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items[$i]
        $numbered += $item
        "{0,3} {1,-13} {2,-25} {3:HH\:mm}-{4:HH\:mm}  {5}" -f ($i+1), $item.Art, $item.Sender, $item.Beginn, $item.Ende, $item.Sendung
    }

    $PrivatSender = @($script:ZattooKanaele.Keys)

    $PrivatNummer = $numbered.Count + 1
    "{0,3} {1,-13} {2}" -f $PrivatNummer, 'Privatsender', '[Sender anzeigen]'

    Write-Host ''
    Write-Host 'Nummer eingeben: Im DVBViewer umschalten. Nur ENTER beendet die Batch.'
    while ($true) {
        $inputValue = Read-Host 'Auswahl'
        if ([string]::IsNullOrWhiteSpace($inputValue)) { break }

        if ($inputValue -notmatch '^\d+$') {
            Write-Host ("Bitte eine Zahl von 1 bis {0} eingeben." -f $PrivatNummer) -ForegroundColor Yellow
            continue
        }
        $selected = [int]$inputValue
        if ($selected -lt 1 -or $selected -gt $PrivatNummer) {
            Write-Host ("Diese Nummer ist nicht vorhanden. Gueltig: 1 bis {0}" -f $PrivatNummer) -ForegroundColor Yellow
            continue
        }

        if ($selected -eq $PrivatNummer) {

            while ($true) {
                Write-Host ''
                Write-Host 'PRIVATSENDER' -ForegroundColor Cyan
                Write-Host '-------------'

                for ($p = 0; $p -lt $PrivatSender.Count; $p++) {
                    Write-Host (" {0} {1}" -f ($p + 1), $PrivatSender[$p])
                }

                Write-Host ' 0 Zurueck'
                Write-Host ''

                $PrivatAuswahl = Read-Host 'Auswahl'

                if ($PrivatAuswahl -eq '0') { break }

                if ($PrivatAuswahl -notmatch '^\d+$') {
                    Write-Host ("Bitte 0 bis {0} eingeben." -f $PrivatSender.Count) -ForegroundColor Yellow
                    continue
                }

                $PrivatIndex = [int]$PrivatAuswahl

                if ($PrivatIndex -lt 1 -or $PrivatIndex -gt $PrivatSender.Count) {
                    Write-Host ("Bitte 0 bis {0} eingeben." -f $PrivatSender.Count) -ForegroundColor Yellow
                    continue
                }

                $PrivatName = $PrivatSender[$PrivatIndex - 1]

                if ($PrivatName -in $script:AktiveZattooSender) {
                    Start-ZattooSender -SenderName $PrivatName
                }
                else {
                    Write-Host ("Testauswahl: {0} - noch kein Senderstart." -f $PrivatName) -ForegroundColor Green
                }
            }

            Write-Host ''
            Write-Host 'AKTUELLES TV-PROGRAMM' -ForegroundColor Cyan

            for ($i = 0; $i -lt $numbered.Count; $i++) {
                $item = $numbered[$i]
                "{0,3} {1,-13} {2,-25} {3:HH\:mm}-{4:HH\:mm}  {5}" -f ($i+1), $item.Art, $item.Sender, $item.Beginn, $item.Ende, $item.Sendung
            }

            "{0,3} {1,-13} {2}" -f $PrivatNummer, 'Privatsender', '[Sender anzeigen]'

            Write-Host ''
            continue
        }

        $chosen = $numbered[$selected - 1]
        Write-Host ("{0}: {1}" -f $chosen.Sender, $chosen.Sendung) -ForegroundColor Cyan
        if ($chosen.Sender -eq 'NTV') {
            $Edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
            $NtvUrl = 'https://zattoo.com/channels?channel=ntv_de'

            Write-Host 'NTV wird ueber Zattoo geoeffnet ...' -ForegroundColor Green

            Start-Process -FilePath $Edge -ArgumentList @(
                '--new-window'
                $NtvUrl
            )

            Start-Sleep -Seconds 5

            Add-Type -AssemblyName System.Windows.Forms

            if (-not ('ZattooFenster' -as [type])) {
                Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class ZattooFenster
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(
        IntPtr hWnd,
        int nCmdShow
    );

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags
    );
}
'@
            }

            $Treffer = New-Object System.Collections.Generic.List[object]

            $Callback = [ZattooFenster+EnumWindowsProc]{
                param($hWnd, $lParam)

                if ([ZattooFenster]::IsWindowVisible($hWnd)) {
                    $Text = New-Object System.Text.StringBuilder 512
                    [void][ZattooFenster]::GetWindowText($hWnd, $Text, $Text.Capacity)

                    if ($Text.ToString() -like 'Sender*Microsoft*Edge*') {
                        $Treffer.Add([pscustomobject]@{ Handle = $hWnd })
                    }
                }

                return $true
            }

            [void][ZattooFenster]::EnumWindows($Callback, [IntPtr]::Zero)

            $Fenster = $Treffer | Select-Object -First 1

            if ($Fenster) {
                $Screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea

                $Links  = [int]($Screen.Width  * 0.731)
                $Oben   = $Screen.Top
                $Breite = [int]($Screen.Width  * 0.269)
                $Hoehe  = [int]($Screen.Height * 0.522)

                [void][ZattooFenster]::ShowWindow(
                    [IntPtr]$Fenster.Handle,
                    9
                )

                Start-Sleep -Milliseconds 500

                [void][ZattooFenster]::SetWindowPos(
                    [IntPtr]$Fenster.Handle,
                    [IntPtr]::Zero,
                    $Links,
                    $Oben,
                    $Breite,
                    $Hoehe,
                    0x0040
                )
            }
        }
        else {
            Start-DVBViewerChannel $chosen.Sender
        }
        Write-Host 'Naechste Nummer eingeben oder nur ENTER zum Beenden.'
    }
}
catch {
    $_ | Out-File -LiteralPath $LogPath -Encoding UTF8
    Write-Host ''
    Write-Host 'FEHLER:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ''
    Write-Host ("Einzelheiten: {0}" -f $LogPath)

    if (-not $ExportOnly) {
        Read-Host 'ENTER beendet'
    }

    exit 1
}










